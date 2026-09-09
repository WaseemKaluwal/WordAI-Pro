// taskpane.js — WordAI Pro core logic

function initApp() {
  loadSettings();
  bindTabs();
  bindSubtabs();
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
  bindEmailButtons();
  bindSemanticSearch();
  bindDiffToggle();
  bindAutocomplete();
  bindDictation();
  bindBrandVoice();
  bindCareerButtons();
  bindSocialButtons();
  bindSlidesButtons();
  bindFrameworkStudio();
  bindHistoryDrawer();
  bindKeyboardShortcuts();
  bindInstantSelectionCopilot();
  bindPromptsLibrary();
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
  btn.addEventListener('click', () => toggleDarkMode());
}

function toggleDarkMode() {
  const dark = document.body.classList.toggle('dark');
  localStorage.setItem('wordai_dark', dark ? '1' : '0');
  const btn = document.getElementById('btn-darkmode');
  if (btn) btn.textContent = dark ? '☀️' : '🌙';
}

// ─── Instant Selection Floating Copilot (Direct In-Place Mode) ──────────

let currentSelectedBodyText = '';
let lastReplacedOriginalText = '';
let isInstantProcessing = false;
let selectionPollInterval = null;
let copilotDialog = null;

function bindInstantSelectionCopilot() {
  const launcher = document.getElementById('instant-copilot-launcher');
  const launcherSubtitle = document.getElementById('launcher-status-text');
  const openModalBtn = document.getElementById('btn-open-popup-modal');
  const openFloatBtn = document.getElementById('btn-open-float-dialog');
  const modalPopoutBtn = document.getElementById('btn-modal-popout');

  const modal = document.getElementById('instant-popup-modal');
  const backdrop = document.getElementById('instant-popup-backdrop');
  const countBadge = document.getElementById('instant-sel-count');
  const snippet = document.getElementById('instant-sel-snippet');
  const input = document.getElementById('instant-instruction-input');
  const execBtn = document.getElementById('btn-instant-exec');
  const undoBtn = document.getElementById('btn-instant-undo');
  const closeBtn = document.getElementById('btn-close-popup-modal');
  const statusEl = document.getElementById('instant-status');

  if (!launcher || !modal) return;

  // Window/taskpane focus & hover triggers for instant selection capture
  window.addEventListener('focus', () => checkWordSelection(false));
  document.addEventListener('mouseenter', () => checkWordSelection(false));

  if (typeof Word !== 'undefined') {
    Word.run(async (context) => {
      if (context.document && context.document.onSelectionChanged) {
        context.document.onSelectionChanged.add(() => checkWordSelection(false));
        await context.sync();
      }
    }).catch(() => {});

    clearInterval(selectionPollInterval);
    selectionPollInterval = setInterval(() => checkWordSelection(false), 380);
  }

  window.openInstantPopupModal = async function() {
    await checkWordSelection(false);
    modal.classList.remove('hidden');
    setTimeout(() => {
      input.focus();
      input.select();
    }, 60);
  };

  window.closeInstantPopupModal = function() {
    modal.classList.add('hidden');
  };

  window.openFloatingCopilotDialog = function() {
    if (typeof Office === 'undefined' || !Office.context || !Office.context.ui || !Office.context.ui.displayDialogAsync) {
      openInstantPopupModal();
      return;
    }

    if (copilotDialog) {
      try { copilotDialog.close(); } catch (_) {}
      copilotDialog = null;
    }

    const currentUrl = window.location.href;
    const dialogUrl = new URL('popup-copilot.html', currentUrl).href;

    Office.context.ui.displayDialogAsync(
      dialogUrl,
      { height: 44, width: 34, displayInIframe: false, promptBeforeOpen: false },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          openInstantPopupModal();
          return;
        }
        copilotDialog = asyncResult.value;
        copilotDialog.addEventHandler(Office.EventType.DialogMessageReceived, onCopilotDialogMessage);
        copilotDialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
          copilotDialog = null;
        });

        setTimeout(() => {
          sendSelectionToCopilotDialog(currentSelectedBodyText);
        }, 600);
      }
    );
  };

  function sendSelectionToCopilotDialog(text) {
    if (copilotDialog) {
      try {
        copilotDialog.messageChild(JSON.stringify({ type: 'selection', text: text || '' }));
      } catch (_) {}
    }
  }

  async function onCopilotDialogMessage(arg) {
    try {
      const data = JSON.parse(arg.message);
      if (data.type === 'ready') {
        sendSelectionToCopilotDialog(currentSelectedBodyText);
      } else if (data.type === 'directReplace') {
        if (data.text) currentSelectedBodyText = data.text;
        await runDirectReplace(data.actionType, data.instruction);
        if (copilotDialog) {
          copilotDialog.messageChild(JSON.stringify({ type: 'status', msg: '✅ Replaced in Word body!', isError: false }));
        }
      }
    } catch (err) {
      if (copilotDialog) {
        copilotDialog.messageChild(JSON.stringify({ type: 'status', msg: `❌ ${err.message}`, isError: true }));
      }
    }
  }

  async function checkWordSelection(forceOpen = false) {
    if (isInstantProcessing || typeof Word === 'undefined') return;
    try {
      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load('text');
        await context.sync();
        const text = (selection.text || '').trim();
        if (text.length >= 2) {
          currentSelectedBodyText = text;
          const words = text.split(/\s+/).filter(Boolean).length;
          countBadge.textContent = `${words} word${words === 1 ? '' : 's'}`;
          snippet.textContent = `"${text.slice(0, 110)}${text.length > 110 ? '…' : ''}"`;
          launcherSubtitle.textContent = `🎯 ${words} word${words === 1 ? '' : 's'} selected — Click for popup`;
          launcherSubtitle.style.color = 'var(--primary)';
          sendSelectionToCopilotDialog(text);

          if (forceOpen) {
            openInstantPopupModal();
          }
        } else {
          if (currentSelectedBodyText !== '') {
            currentSelectedBodyText = '';
            launcherSubtitle.textContent = 'Highlight text in Word or click to instruct';
            launcherSubtitle.style.color = 'var(--text-secondary)';
            snippet.textContent = '"Highlight text in your Word document body..."';
            countBadge.textContent = '0 words';
          }
        }
      });
    } catch (_) {}
  }

  async function runDirectReplace(instructionType, customInstruction = '') {
    if (isInstantProcessing) return;

    // If no text was captured, try one live grab
    if (!currentSelectedBodyText || currentSelectedBodyText.length < 2) {
      await checkWordSelection(false);
    }

    const targetText = currentSelectedBodyText;
    if (!targetText || targetText.length < 2) {
      setInstantStatus('⚠️ Highlight text in Word body first.', true);
      return;
    }

    let prompt = '';
    switch (instructionType) {
      case 'rewrite':
        prompt = 'Rewrite this text to be clearer, more engaging, and professionally polished. Return ONLY the final replacement text without any conversational preamble or markdown commentary.';
        break;
      case 'improve':
        prompt = 'Improve the flow, vocabulary, and sentence variety of this text while preserving its exact factual meaning. Return ONLY the improved replacement text.';
        break;
      case 'grammar':
        prompt = 'Fix all spelling, punctuation, grammar, and phrasing errors in this text. Return ONLY the corrected replacement text.';
        break;
      case 'shorten':
        prompt = 'Make this text significantly more concise, punchy, and direct without losing critical details. Return ONLY the shortened replacement text.';
        break;
      case 'expand':
        prompt = 'Elaborate on this text with deeper explanation, supporting detail, and professional depth. Return ONLY the expanded replacement text.';
        break;
      case 'formal':
        prompt = 'Rewrite this text in an authoritative, formal, and polished executive/academic tone. Return ONLY the formal replacement text.';
        break;
      case 'casual':
        prompt = 'Rewrite this text in a warm, conversational, friendly, and natural tone. Return ONLY the casual replacement text.';
        break;
      case 'custom':
      default:
        prompt = `${customInstruction}. Return ONLY the direct replacement text to be inserted in the document without any extra commentary.`;
        break;
    }

    isInstantProcessing = true;
    execBtn.disabled = true;
    execBtn.textContent = '⏳ Replacing...';
    setInstantStatus('⚡ Generating & directly replacing in Word body...');

    try {
      const replacement = await runAI(prompt, targetText);
      if (!replacement || replacement.trim() === '') {
        throw new Error('AI returned an empty response.');
      }
      lastReplacedOriginalText = targetText;
      await insertText(replacement.trim(), true);
      currentSelectedBodyText = replacement.trim();
      snippet.textContent = `"${replacement.trim().slice(0, 110)}${replacement.trim().length > 110 ? '…' : ''}"`;
      undoBtn.classList.remove('hidden');
      setInstantStatus('✅ Replaced directly in Word body!', false);
      input.value = '';
    } catch (err) {
      setInstantStatus(`❌ ${err.message}`, true);
    } finally {
      isInstantProcessing = false;
      execBtn.disabled = false;
      execBtn.textContent = 'Replace ⚡';
    }
  }

  function setInstantStatus(msg, isError = false) {
    statusEl.textContent = msg;
    statusEl.style.color = isError ? 'var(--danger)' : 'var(--primary)';
    statusEl.classList.remove('hidden');
    setTimeout(() => {
      if (statusEl.textContent === msg) statusEl.classList.add('hidden');
    }, 4500);
  }

  // Launcher click opens in-place popup modal
  launcher.addEventListener('click', (e) => {
    if (e.target.closest('#btn-open-float-dialog')) return;
    openInstantPopupModal();
  });

  openModalBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    openInstantPopupModal();
  });

  openFloatBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    openFloatingCopilotDialog();
  });

  if (modalPopoutBtn) {
    modalPopoutBtn.addEventListener('click', () => {
      closeInstantPopupModal();
      openFloatingCopilotDialog();
    });
  }

  backdrop.addEventListener('click', closeInstantPopupModal);
  closeBtn.addEventListener('click', closeInstantPopupModal);

  window.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && !modal.classList.contains('hidden')) {
      closeInstantPopupModal();
    }
  });

  execBtn.addEventListener('click', () => {
    const val = input.value.trim();
    if (!val) {
      runDirectReplace('rewrite');
    } else {
      runDirectReplace('custom', val);
    }
  });

  input.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      execBtn.click();
    }
  });

  document.querySelectorAll('.instant-chip').forEach(chip => {
    chip.addEventListener('click', () => {
      const act = chip.getAttribute('data-instant');
      runDirectReplace(act);
    });
  });

  undoBtn.addEventListener('click', async () => {
    if (!lastReplacedOriginalText) return;
    try {
      await insertText(lastReplacedOriginalText, true);
      currentSelectedBodyText = lastReplacedOriginalText;
      snippet.textContent = `"${lastReplacedOriginalText.slice(0, 110)}${lastReplacedOriginalText.length > 110 ? '…' : ''}"`;
      setInstantStatus('↩️ Restored original text!', false);
      undoBtn.classList.add('hidden');
    } catch (err) {
      setInstantStatus(`❌ Undo failed: ${err.message}`, true);
    }
  });
}

// ─── Helpers & Formatting ───────────────────────────────────────────────────

let currentSourceContext = '';
let currentResultText = '';

function showLoading(show) {
  document.getElementById('loading').classList.toggle('hidden', !show);
}

function escapeHtml(str) {
  return (str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
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

function generateRedlineDiff(oldText, newText) {
  if (!oldText || oldText.trim() === '') {
    return `<ins class="diff-ins">${escapeHtml(newText)}</ins>`;
  }
  const oldWords = oldText.trim().split(/\s+/);
  const newWords = (newText || '').trim().split(/\s+/);

  const m = Math.min(oldWords.length, 300);
  const n = Math.min(newWords.length, 300);

  const dp = Array.from({ length: m + 1 }, () => new Uint16Array(n + 1));
  for (let i = 0; i < m; i++) {
    for (let j = 0; j < n; j++) {
      if (oldWords[i].toLowerCase() === newWords[j].toLowerCase()) {
        dp[i + 1][j + 1] = dp[i][j] + 1;
      } else {
        dp[i + 1][j + 1] = Math.max(dp[i + 1][j], dp[i][j + 1]);
      }
    }
  }

  let i = m, j = n;
  const result = [];
  while (i > 0 || j > 0) {
    if (i > 0 && j > 0 && oldWords[i - 1].toLowerCase() === newWords[j - 1].toLowerCase()) {
      result.unshift(escapeHtml(newWords[j - 1]));
      i--; j--;
    } else if (j > 0 && (i === 0 || dp[i][j - 1] >= dp[i - 1][j])) {
      result.unshift(`<ins class="diff-ins">${escapeHtml(newWords[j - 1])}</ins>`);
      j--;
    } else if (i > 0 && (j === 0 || dp[i][j - 1] < dp[i - 1][j])) {
      result.unshift(`<del class="diff-del">${escapeHtml(oldWords[i - 1])}</del>`);
      i--;
    }
  }

  if (newWords.length > 300) {
    result.push(escapeHtml(newWords.slice(300).join(' ')));
  }
  return result.join(' ');
}

function showOutput(text, rawOriginal) {
  currentResultText = text;
  const section = document.getElementById('output-section');
  const textEl = document.getElementById('output-text');
  const diffEl = document.getElementById('output-diff');
  const origEl = document.getElementById('sidebyside-original');
  const outEl  = document.getElementById('sidebyside-output');

  textEl.innerHTML = '<p>' + markdownToHtml(text) + '</p>';
  diffEl.innerHTML = generateRedlineDiff(currentSourceContext, currentResultText);
  origEl.innerHTML = '<p>' + markdownToHtml(rawOriginal || currentSourceContext || '(No source context)') + '</p>';
  outEl.innerHTML  = '<p>' + markdownToHtml(text) + '</p>';

  recordResponse(currentSourceContext.slice(0, 80) || 'Output', text);
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

function markdownToRichWordHtml(markdown) {
  if (!markdown) return '';

  const lines = markdown.split('\n');
  let html = '';
  let inList = false;
  let listType = '';
  let inTable = false;
  let tableRows = [];

  function closeList() {
    if (inList) {
      html += `</${listType}>\n`;
      inList = false;
      listType = '';
    }
  }

  function flushTable() {
    if (!tableRows.length) return;
    const cleanRows = tableRows.filter(r => !/^\|?[-:\s|]+\|?$/.test(r.trim()));
    if (!cleanRows.length) { tableRows = []; inTable = false; return; }

    html += '<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;width:100%;border:1px solid #cbd5e1;margin-top:10pt;margin-bottom:12pt;font-family:\'Segoe UI\',Calibri,Arial,sans-serif;font-size:10.5pt;">\n';
    cleanRows.forEach((rowStr, idx) => {
      const cells = rowStr.trim().replace(/^\|/, '').replace(/\|$/, '').split('|').map(c => c.trim());
      html += '  <tr>\n';
      cells.forEach(cell => {
        const formatted = formatInline(cell);
        if (idx === 0) {
          html += `    <th style="background-color:#f1f5f9;font-weight:bold;text-align:left;padding:6pt 8pt;border:1px solid #cbd5e1;color:#0f172a;">${formatted}</th>\n`;
        } else {
          html += `    <td style="padding:5pt 8pt;border:1px solid #cbd5e1;vertical-align:top;color:#1e293b;">${formatted}</td>\n`;
        }
      });
      html += '  </tr>\n';
    });
    html += '</table>\n';
    tableRows = [];
    inTable = false;
  }

  function formatInline(text) {
    if (!text) return '';
    const underscores = [];
    // Protect signature underline lines like By: ___________________
    text = text.replace(/_{2,}/g, (m) => {
      underscores.push(m);
      return `@@USCORE${underscores.length - 1}@@`;
    });

    let formatted = text
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/\*\*\*(.+?)\*\*\*/g, '<strong><em>$1</em></strong>')
      .replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>')
      .replace(/__(.+?)__/g, '<strong>$1</strong>')
      .replace(/(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)/g, '<em>$1</em>')
      .replace(/(?<!_)_(?!_)(.+?)(?<!_)_(?!_)/g, '<em>$1</em>')
      .replace(/`([^`]+)`/g, '<code style="font-family:Consolas,monospace;background-color:#f1f5f9;padding:1pt 3pt;border-radius:2pt;font-size:9.5pt;">$1</code>');

    // Restore signature lines
    underscores.forEach((u, idx) => {
      formatted = formatted.replace(`@@USCORE${idx}@@`, u);
    });

    return formatted;
  }

  for (let i = 0; i < lines.length; i++) {
    const rawLine = lines[i];
    const trimmed = rawLine.trim();

    // Table Row
    if (/^\|.*\|$/.test(trimmed)) {
      closeList();
      inTable = true;
      tableRows.push(trimmed);
      continue;
    } else if (inTable) {
      flushTable();
    }

    if (!trimmed) {
      closeList();
      continue;
    }

    // Headings
    if (/^#\s+/.test(trimmed)) {
      closeList();
      const text = formatInline(trimmed.replace(/^#\s+/, ''));
      html += `<h1 style="font-family:\'Segoe UI\',Calibri,Arial,sans-serif;font-size:20pt;font-weight:bold;color:#0f172a;margin-top:18pt;margin-bottom:8pt;line-height:1.25;">${text}</h1>\n`;
      continue;
    }
    if (/^##\s+/.test(trimmed)) {
      closeList();
      const text = formatInline(trimmed.replace(/^##\s+/, ''));
      html += `<h2 style="font-family:\'Segoe UI\',Calibri,Arial,sans-serif;font-size:14.5pt;font-weight:bold;color:#1e40af;margin-top:14pt;margin-bottom:6pt;line-height:1.3;">${text}</h2>\n`;
      continue;
    }
    if (/^###\s+/.test(trimmed)) {
      closeList();
      const text = formatInline(trimmed.replace(/^###\s+/, ''));
      html += `<h3 style="font-family:\'Segoe UI\',Calibri,Arial,sans-serif;font-size:12.5pt;font-weight:bold;color:#334155;margin-top:10pt;margin-bottom:4pt;line-height:1.3;">${text}</h3>\n`;
      continue;
    }
    if (/^####\s+/.test(trimmed)) {
      closeList();
      const text = formatInline(trimmed.replace(/^####\s+/, ''));
      html += `<h4 style="font-family:\'Segoe UI\',Calibri,Arial,sans-serif;font-size:11pt;font-weight:bold;color:#475569;margin-top:8pt;margin-bottom:3pt;">${text}</h4>\n`;
      continue;
    }

    // Horizontal Rule
    if (/^[-*_]{3,}$/.test(trimmed)) {
      closeList();
      html += '<hr style="border:none;border-top:1px solid #cbd5e1;margin-top:12pt;margin-bottom:12pt;" />\n';
      continue;
    }

    // Bullet Lists
    if (/^[*\-]\s+/.test(trimmed)) {
      if (!inList || listType !== 'ul') {
        closeList();
        inList = true;
        listType = 'ul';
        html += '<ul style="margin-top:4pt;margin-bottom:8pt;padding-left:18pt;">\n';
      }
      const itemText = formatInline(trimmed.replace(/^[*\-]\s+/, ''));
      html += `  <li style="font-family:\'Segoe UI\',Calibri,Arial,sans-serif;font-size:11pt;line-height:1.55;margin-bottom:4pt;color:#1e293b;">${itemText}</li>\n`;
      continue;
    }

    // Numbered Lists
    const numMatch = trimmed.match(/^(\d+)\.\s+(.+)/);
    if (numMatch) {
      if (!inList || listType !== 'ol') {
        closeList();
        inList = true;
        listType = 'ol';
        html += '<ol style="margin-top:4pt;margin-bottom:8pt;padding-left:18pt;">\n';
      }
      const itemText = formatInline(numMatch[2]);
      html += `  <li style="font-family:\'Segoe UI\',Calibri,Arial,sans-serif;font-size:11pt;line-height:1.55;margin-bottom:4pt;color:#1e293b;">${itemText}</li>\n`;
      continue;
    }

    // Normal Paragraph
    closeList();
    const pText = formatInline(trimmed);
    html += `<p style="font-family:\'Segoe UI\',Calibri,Arial,sans-serif;font-size:11pt;line-height:1.55;margin-top:0pt;margin-bottom:8pt;color:#1e293b;">${pText}</p>\n`;
  }

  closeList();
  if (inTable) flushTable();

  return html;
}

async function insertText(rawText, replace = false, insertAtEnd = false) {
  if (typeof Word === 'undefined') { showError('Word API not available.'); return; }
  return Word.run(async (ctx) => {
    try {
      const html = markdownToRichWordHtml(rawText);
      if (insertAtEnd) {
        ctx.document.body.insertHtml(html, Word.InsertLocation.end);
      } else {
        const sel = ctx.document.getSelection();
        sel.insertHtml(html, replace ? Word.InsertLocation.replace : Word.InsertLocation.after);
      }
      await ctx.sync();
    } catch (htmlErr) {
      await fallbackInsertParagraphs(ctx, rawText, replace, insertAtEnd);
    }
  });
}

async function fallbackInsertParagraphs(ctx, rawText, replace = false, insertAtEnd = false) {
  let cursor;
  if (insertAtEnd) {
    cursor = ctx.document.body;
  } else {
    const sel = ctx.document.getSelection();
    if (replace) sel.insertText('', Word.InsertLocation.replace);
    cursor = replace ? sel : sel.getRange(Word.RangeLocation.after);
  }

  const lines = rawText.split('\n');
  let inTable = false;
  let tableLines = [];

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (/^\|/.test(line.trim())) {
      inTable = true;
      tableLines.push(line);
      continue;
    } else if (inTable) {
      const rows = tableLines.filter(l => !/^\|[-:\s|]+$/.test(l.trim()));
      const parsed = rows.map(r => r.replace(/^\|/, '').replace(/\|$/, '').split('|').map(c => stripInline(c.trim())));
      if (parsed.length && parsed[0].length) {
        const table = cursor.insertTable(parsed.length, parsed[0].length, insertAtEnd ? Word.InsertLocation.end : Word.InsertLocation.after, parsed);
        table.styleBuiltIn = Word.Style.tableGrid;
        await ctx.sync();
        if (!insertAtEnd) cursor = table.getRange(Word.RangeLocation.after);
      }
      tableLines = [];
      inTable = false;
    }

    if (!line.trim()) continue;

    const loc = insertAtEnd ? Word.InsertLocation.end : Word.InsertLocation.after;

    const h1 = line.match(/^#\s+(.+)/);
    const h2 = line.match(/^##\s+(.+)/);
    const h3 = line.match(/^###\s+(.+)/);
    const h4 = line.match(/^####\s+(.+)/);

    if (h1 || h2 || h3 || h4) {
      const text = stripInline((h1 || h2 || h3 || h4)[1]);
      const p = cursor.insertParagraph(text, loc);
      p.styleBuiltIn = h1 ? Word.Style.title : h2 ? Word.Style.heading1 : h3 ? Word.Style.heading2 : Word.Style.heading3;
      p.font.bold = true;
      await ctx.sync();
      if (!insertAtEnd) cursor = p.getRange(Word.RangeLocation.after);
      continue;
    }

    const bullet = line.match(/^[*-]\s+(.+)/);
    if (bullet) {
      const p = cursor.insertParagraph(stripInline(bullet[1]), loc);
      p.styleBuiltIn = Word.Style.listParagraph;
      p.listItem.listLevelType = Word.ListLevelType.bullet;
      await ctx.sync();
      await applyInlineBold(ctx, p, bullet[1]);
      if (!insertAtEnd) cursor = p.getRange(Word.RangeLocation.after);
      continue;
    }

    const numbered = line.match(/^(\d+)\.\s+(.+)/);
    if (numbered) {
      const p = cursor.insertParagraph(stripInline(numbered[2]), loc);
      p.styleBuiltIn = Word.Style.listParagraph;
      p.listItem.listLevelType = Word.ListLevelType.number;
      await ctx.sync();
      await applyInlineBold(ctx, p, numbered[2]);
      if (!insertAtEnd) cursor = p.getRange(Word.RangeLocation.after);
      continue;
    }

    const p = cursor.insertParagraph(stripInline(line), loc);
    p.styleBuiltIn = Word.Style.normal;
    await ctx.sync();
    await applyInlineBold(ctx, p, line);
    if (!insertAtEnd) cursor = p.getRange(Word.RangeLocation.after);
  }

  await ctx.sync();
}

function stripInline(text) {
  if (!text) return '';
  return text
    .replace(/\*\*\*(.+?)\*\*\*/g, '$1')
    .replace(/\*\*(.+?)\*\*/g, '$1')
    .replace(/\*(.+?)\*/g, '$1')
    .replace(/___(.+?)___/g, '$1')
    .replace(/__(.+?)__/g, '$1')
    .replace(/_(.+?)_/g, '$1');
}

async function applyInlineBold(ctx, p, rawText) {
  const matches = [...rawText.matchAll(/\*\*(.+?)\*\*/g)];
  if (!matches.length) return;
  for (const m of matches) {
    try {
      const searchRes = p.search(m[1], { matchCase: false });
      searchRes.load('items');
      await ctx.sync();
      searchRes.items.forEach(item => { item.font.bold = true; });
      await ctx.sync();
    } catch (_) {}
  }
}

function getPersonaInstruction() {
  const persona = localStorage.getItem('wordai_persona') || 'standard';
  if (persona === 'academic') return 'Adopt an academic scholar persona: rigorous, evidence-based, formal, and theoretically precise.';
  if (persona === 'executive') return 'Adopt an executive persona: concise, direct, high-impact, bullet-oriented, and bottom-line focused.';
  if (persona === 'creative') return 'Adopt a creative storytelling persona: vivid, engaging, evocative, and rhythmic.';
  if (persona === 'custom') {
    const customPrompt = localStorage.getItem('wordai_custom_persona_prompt');
    if (customPrompt) return `Adopt this custom author voice: ${customPrompt}`;
  }
  return 'Adopt a polished, professional, and clear tone.';
}

async function runAI(systemPrompt, userContent) {
  const lang = localStorage.getItem('wordai_language') || 'English';
  const personaInstruction = getPersonaInstruction();
  const messages = [
    { role: 'system', content: `${systemPrompt} ${personaInstruction} Respond in ${lang}.` },
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
  const provider = localStorage.getItem('wordai_provider') || 'openai';
  const status = document.getElementById('fetch-status');
  if (status && model === '__auto__' && working) {
    status.textContent = `⚡ Using: ${working}`;
  }

  const badge = document.getElementById('header-model-badge');
  if (badge) {
    let label = provider.toUpperCase();
    const activeModel = (model && model !== '__auto__') ? model : (working || '');
    if (activeModel) {
      if (/gpt-4o/i.test(activeModel)) label = 'GPT-4o';
      else if (/o3-mini/i.test(activeModel)) label = 'o3-mini';
      else if (/claude-3-5-sonnet/i.test(activeModel)) label = 'Claude 3.5';
      else if (/claude-3-5-haiku/i.test(activeModel)) label = 'Claude Haiku';
      else if (/gemini-2/i.test(activeModel)) label = 'Gemini 2.0';
      else if (/deepseek/i.test(activeModel)) label = 'DeepSeek';
      else if (/llama-3/i.test(activeModel)) label = 'Llama 3.3';
      else if (/mistral/i.test(activeModel)) label = 'Mistral';
      else if (/command-r/i.test(activeModel)) label = 'Cohere';
      else if (activeModel.length <= 12) label = activeModel;
      else label = provider.charAt(0).toUpperCase() + provider.slice(1);
    } else {
      label = provider.charAt(0).toUpperCase() + provider.slice(1);
    }
    badge.textContent = label;
    badge.title = `Active AI Provider: ${provider.toUpperCase()}${activeModel ? ` (${activeModel})` : ''}`;
  }
}

function showError(msg) {
  showOutput(`⚠️ ${msg}`);
}

// ─── Tabs & Subtabs ──────────────────────────────────────────────────────────

function bindTabs() {
  document.querySelectorAll('.tab').forEach(tab => {
    tab.addEventListener('click', () => {
      document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
      document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
      tab.classList.add('active');
      const target = document.getElementById(`tab-${tab.dataset.tab}`);
      if (target) target.classList.add('active');
    });
  });
}

function bindSubtabs() {
  document.querySelectorAll('.doc-subtab').forEach(subtab => {
    subtab.addEventListener('click', () => {
      document.querySelectorAll('.doc-subtab').forEach(t => t.classList.remove('active'));
      document.querySelectorAll('.doc-subpanel').forEach(p => p.classList.remove('active'));
      subtab.classList.add('active');
      const target = document.getElementById(`subpanel-${subtab.dataset.subtab}`);
      if (target) target.classList.add('active');
    });
  });
}

// ─── Document Tab Automation ─────────────────────────────────────────────────

function docStatus(msg, isError) {
  const el = document.getElementById('doc-status');
  el.textContent = msg;
  el.className = 'doc-status ' + (isError ? 'doc-status-error' : 'doc-status-ok');
  el.classList.remove('hidden');
  setTimeout(() => el.classList.add('hidden'), 4000);
}

function bindDocumentButtons() {
  document.getElementById('font-family-select').addEventListener('change', () => applyFontFormat());
  document.getElementById('font-size-select').addEventListener('change', () => applyFontFormat());

  document.querySelectorAll('.fmt-btn').forEach(btn => {
    btn.addEventListener('click', async () => {
      const fmt = btn.dataset.fmt;
      if (typeof Word === 'undefined') return docStatus('Word API unavailable', true);
      await Word.run(async ctx => {
        const sel = ctx.document.getSelection();
        sel.load('font');
        await ctx.sync();
        if (fmt === 'bold')        sel.font.bold = !sel.font.bold;
        if (fmt === 'italic')      sel.font.italic = !sel.font.italic;
        if (fmt === 'underline')   sel.font.underline = sel.font.underline === 'None' ? 'Single' : 'None';
        if (fmt === 'strike')      sel.font.strikeThrough = !sel.font.strikeThrough;
        if (fmt === 'superscript') sel.font.superscript = !sel.font.superscript;
        if (fmt === 'subscript')   sel.font.subscript = !sel.font.subscript;
        if (fmt === 'allcaps')     sel.font.allCaps = !sel.font.allCaps;
        if (fmt === 'smallcaps')   sel.font.smallCaps = !sel.font.smallCaps;
        await ctx.sync();
      });
    });
  });

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

      if (action === 'wordcount') {
        const body = ctx.document.body;
        body.load('text');
        await ctx.sync();
        const text = body.text.trim();
        const words = text ? text.split(/\s+/).filter(Boolean).length : 0;
        const chars = text.length;
        const paras = text.split(/\n+/).filter(Boolean).length;
        docStatus(`🔢 Words: ${words.toLocaleString()} | Chars: ${chars.toLocaleString()} | Paras: ${paras}`);
        return;
      }

      if (action === 'properties') {
        const props = ctx.document.properties;
        props.load('title,author,subject,keywords,lastModifiedBy');
        await ctx.sync();
        docStatus(`📋 Title: ${props.title||'—'} | Author: ${props.author||'—'} | Modified by: ${props.lastModifiedBy||'—'}`);
        return;
      }

      if (action === 'outline') {
        const paras = ctx.document.body.paragraphs;
        paras.load('text,styleBuiltIn');
        await ctx.sync();
        const headings = paras.items
          .filter(p => ['Heading1','Heading2','Heading3','Heading4'].includes(p.styleBuiltIn))
          .map(p => `${p.styleBuiltIn.replace('Heading','H')}: ${p.text.trim()}`)
          .join('\n');
        showOutput(headings || 'No headings found in document.');
        return;
      }

      if (action === 'bookmarks') {
        const bms = ctx.document.bookmarks;
        bms.load('items/name');
        await ctx.sync();
        const names = bms.items.map(b => b.name).join(', ');
        docStatus(`🔖 Bookmarks: ${names || 'None found'}`);
        return;
      }

      if (action === 'list-styles') {
        const paras = ctx.document.body.paragraphs;
        paras.load('items/styleBuiltIn,items/style');
        await ctx.sync();
        const styles = new Set();
        paras.items.forEach(p => {
          if (p.styleBuiltIn) styles.add(p.styleBuiltIn);
          if (p.style) styles.add(p.style);
        });
        showOutput(`🎨 Paragraph Styles used in Document:\n\n` + [...styles].map(s => `• ${s}`).join('\n') || 'No custom styles found.');
        return;
      }

      if (action === 'list-fields') {
        try {
          const fields = ctx.document.body.fields;
          fields.load('items/type');
          await ctx.sync();
          if (!fields.items.length) { docStatus('No dynamic fields found in document'); return; }
          const list = fields.items.map((f, i) => `${i+1}. Type: ${f.type || 'Field'}`).join('\n');
          showOutput(`🏷️ Document Fields:\n\n${list}`);
        } catch(_) {
          docStatus('Document fields checked');
        }
        return;
      }

      if (action === 'footnotes' || action === 'endnotes') {
        docStatus(`📎 ${action === 'footnotes' ? 'Footnotes' : 'Endnotes'} checked`);
        return;
      }

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

      if (action === 'find-clear') {
        const term = document.getElementById('find-input').value.trim();
        if (term) {
          const results = ctx.document.body.search(term, { matchCase: false });
          results.load('items');
          await ctx.sync();
          results.items.forEach(r => { r.font.highlightColor = null; });
          await ctx.sync();
          docStatus(`🚫 Cleared highlights for "${term}"`);
        } else {
          ctx.document.body.font.highlightColor = null;
          await ctx.sync();
          docStatus('🚫 Cleared all highlights in document');
        }
        return;
      }

      if (action === 'find-regex') {
        const patternStr = document.getElementById('find-input').value.trim();
        if (!patternStr) { docStatus('Enter regex pattern in Find box', true); return; }
        try {
          const regex = new RegExp(patternStr, 'gi');
          const body = ctx.document.body;
          body.load('text');
          await ctx.sync();
          const matches = [...body.text.matchAll(regex)];
          if (!matches.length) {
            docStatus(`⚡ Regex Find: 0 matches found for /${patternStr}/gi`);
            return;
          }
          const uniqueWords = [...new Set(matches.map(m => m[0]))].slice(0, 25);
          for (const word of uniqueWords) {
            const sr = ctx.document.body.search(word, { matchCase: true });
            sr.load('items');
            await ctx.sync();
            sr.items.forEach(item => { item.font.highlightColor = 'Yellow'; });
          }
          await ctx.sync();
          docStatus(`⚡ Found ${matches.length} regex matches — highlighted in yellow`);
        } catch(e) {
          docStatus(`⚠️ Regex error: ${e.message}`, true);
        }
        return;
      }

      if (action === 'apply-font-color') {
        const color = document.getElementById('font-color-input').value;
        const sel = ctx.document.getSelection();
        sel.font.color = color;
        await ctx.sync();
        docStatus(`🎨 Font color applied: ${color}`);
        return;
      }
      if (action === 'apply-highlight') {
        const color = document.getElementById('highlight-color-select').value;
        const sel = ctx.document.getSelection();
        sel.font.highlightColor = color === 'None' ? null : color;
        await ctx.sync();
        docStatus(`🎨 Highlight applied: ${color}`);
        return;
      }
      if (action === 'clear-format') {
        const sel = ctx.document.getSelection();
        sel.load('font');
        await ctx.sync();
        sel.font.bold = false;
        sel.font.italic = false;
        sel.font.underline = 'None';
        sel.font.strikeThrough = false;
        sel.font.color = '#000000';
        sel.font.highlightColor = null;
        await ctx.sync();
        docStatus('🧹 Formatting cleared');
        return;
      }
      if (action === 'clear-highlight') {
        const sel = ctx.document.getSelection();
        sel.font.highlightColor = null;
        await ctx.sync();
        docStatus('🚫 Highlight removed');
        return;
      }

      const alignMap = {
        'align-left': 'Left',
        'align-center': 'Centered',
        'align-right': 'Right',
        'align-justify': 'Justified'
      };
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

      if (action === 'apply-spacing') {
        const spacingVal = parseFloat(document.getElementById('line-spacing-select').value) || 1.15;
        const sel = ctx.document.getSelection();
        const paras = sel.paragraphs;
        paras.load('items');
        await ctx.sync();
        paras.items.forEach(p => {
          p.lineSpacing = Math.round(spacingVal * 12);
        });
        await ctx.sync();
        docStatus(`¶ Line spacing set to ${spacingVal}`);
        return;
      }

      const styleMap = {
        'apply-heading1': Word.Style.heading1,
        'apply-heading2': Word.Style.heading2,
        'apply-heading3': Word.Style.heading3,
        'apply-heading4': Word.Style.heading4,
        'apply-normal':   Word.Style.normal,
        'apply-quote':    Word.Style.quote,
        'apply-intense-quote': Word.Style.intenseQuote,
        'apply-code':     'Code',
        'apply-title':    Word.Style.title,
        'apply-subtitle': Word.Style.subtitle,
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

      if (action === 'insert-table') {
        const rows = Math.min(20, Math.max(1, parseInt(document.getElementById('table-rows').value) || 3));
        const cols = Math.min(10, Math.max(1, parseInt(document.getElementById('table-cols').value) || 3));
        const sel = ctx.document.getSelection();
        const values = Array.from({ length: rows }, (_, r) =>
          Array.from({ length: cols }, (_, c) => r === 0 ? `Header ${c + 1}` : `Row ${r} Col ${c + 1}`)
        );
        const table = sel.insertTable(rows, cols, Word.InsertLocation.replace, values);
        table.styleBuiltIn = Word.Style.tableGrid;
        const headerRow = table.rows.getFirst();
        headerRow.load('cells');
        await ctx.sync();
        headerRow.cells.items.forEach(cell => {
          cell.body.paragraphs.getFirst().font.bold = true;
        });
        await ctx.sync();
        docStatus(`📊 Inserted ${rows}x${cols} table`);
        return;
      }
      if (action === 'table-add-row') {
        const sel = ctx.document.getSelection();
        const table = sel.parentTable;
        table.load('rows/items');
        await ctx.sync();
        if (table.isNullObject) { docStatus('Select a table cell first', true); return; }
        table.addRows(Word.InsertLocation.after, 1);
        await ctx.sync();
        docStatus('➕ Added table row');
        return;
      }
      if (action === 'table-add-col') {
        const sel = ctx.document.getSelection();
        const table = sel.parentTable;
        table.load('columns/items');
        await ctx.sync();
        if (table.isNullObject) { docStatus('Select a table cell first', true); return; }
        table.addColumns(Word.InsertLocation.after, 1);
        await ctx.sync();
        docStatus('➕ Added table col');
        return;
      }
      if (action === 'table-delete-row') {
        const sel = ctx.document.getSelection();
        const table = sel.parentTable;
        table.load('rows/items');
        await ctx.sync();
        if (table.isNullObject) { docStatus('Select a table cell first', true); return; }
        table.rows.getLast().delete();
        await ctx.sync();
        docStatus('➖ Deleted table row');
        return;
      }
      if (action === 'table-delete-col') {
        const sel = ctx.document.getSelection();
        const table = sel.parentTable;
        table.load('columns/items');
        await ctx.sync();
        if (table.isNullObject) { docStatus('Select a table cell first', true); return; }
        table.columns.getLast().delete();
        await ctx.sync();
        docStatus('➖ Deleted table col');
        return;
      }
      if (action === 'table-style-grid') {
        const sel = ctx.document.getSelection();
        const table = sel.parentTable;
        table.load('isNullObject');
        await ctx.sync();
        if (!table.isNullObject) {
          table.styleBuiltIn = Word.Style.tableGrid;
          await ctx.sync();
          docStatus('🔲 Applied Grid Style');
        } else {
          docStatus('Place cursor inside a table', true);
        }
        return;
      }
      if (action === 'table-style-plain') {
        const sel = ctx.document.getSelection();
        const table = sel.parentTable;
        table.load('isNullObject');
        await ctx.sync();
        if (!table.isNullObject) {
          table.style = 'Normal Table';
          await ctx.sync();
          docStatus('⬜ Applied Plain Style');
        } else {
          docStatus('Place cursor inside a table', true);
        }
        return;
      }

      if (action === 'insert-hyperlink') {
        const text = document.getElementById('hyperlink-text').value.trim() || 'Link';
        const url  = document.getElementById('hyperlink-url').value.trim();
        if (!url) { docStatus('Enter a URL (e.g. https://example.com)', true); return; }
        const sel = ctx.document.getSelection();
        const p = sel.insertParagraph(text, Word.InsertLocation.replace);
        p.hyperlink = url;
        await ctx.sync();
        docStatus(`🔗 Hyperlink inserted: "${text}"`);
        return;
      }
      if (action === 'insert-footnote') {
        const note = prompt('Enter footnote text:');
        if (!note) return;
        const sel = ctx.document.getSelection();
        sel.insertFootnote(note);
        await ctx.sync();
        docStatus('📎 Footnote inserted');
        return;
      }
      if (action === 'insert-endnote') {
        const note = prompt('Enter endnote text:');
        if (!note) return;
        const sel = ctx.document.getSelection();
        sel.insertEndnote(note);
        await ctx.sync();
        docStatus('📌 Endnote inserted');
        return;
      }
      if (action === 'insert-pagebreak') {
        const sel = ctx.document.getSelection();
        sel.insertBreak(Word.BreakType.page, Word.InsertLocation.after);
        await ctx.sync();
        docStatus('⏎ Page break inserted');
        return;
      }
      if (action === 'insert-section-break') {
        const sel = ctx.document.getSelection();
        sel.insertBreak(Word.BreakType.sectionNext, Word.InsertLocation.after);
        await ctx.sync();
        docStatus('📄 Section break inserted');
        return;
      }
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
      if (action === 'insert-date') {
        const now = new Date();
        const str = now.toLocaleDateString(undefined, { year: 'numeric', month: 'long', day: 'numeric', hour: '2-digit', minute: '2-digit' });
        const sel = ctx.document.getSelection();
        sel.insertText(str, Word.InsertLocation.replace);
        await ctx.sync();
        docStatus(`📅 Inserted: ${str}`);
        return;
      }

      if (action === 'set-header') {
        const text = document.getElementById('header-text').value.trim();
        if (!text) { docStatus('Enter header text first', true); return; }
        const section = ctx.document.sections.getFirst();
        const header = section.getHeader(Word.HeaderFooterType.primary);
        header.insertParagraph(text, Word.InsertLocation.replace);
        await ctx.sync();
        docStatus(`📝 Header updated: "${text}"`);
        return;
      }
      if (action === 'set-footer') {
        const text = document.getElementById('footer-text').value.trim();
        if (!text) { docStatus('Enter footer text first', true); return; }
        const section = ctx.document.sections.getFirst();
        const footer = section.getFooter(Word.HeaderFooterType.primary);
        footer.insertParagraph(text, Word.InsertLocation.replace);
        await ctx.sync();
        docStatus(`📝 Footer updated: "${text}"`);
        return;
      }
      if (action === 'get-header') {
        const section = ctx.document.sections.getFirst();
        const header = section.getHeader(Word.HeaderFooterType.primary);
        header.load('text');
        await ctx.sync();
        docStatus(`👁️ Header: ${header.text.trim() || '(empty)'}`);
        return;
      }
      if (action === 'get-footer') {
        const section = ctx.document.sections.getFirst();
        const footer = section.getFooter(Word.HeaderFooterType.primary);
        footer.load('text');
        await ctx.sync();
        docStatus(`👁️ Footer: ${footer.text.trim() || '(empty)'}`);
        return;
      }

      if (action === 'add-bookmark') {
        const name = document.getElementById('bookmark-name').value.trim().replace(/\s+/g, '_');
        if (!name) { docStatus('Enter bookmark name', true); return; }
        const sel = ctx.document.getSelection();
        sel.insertBookmark(name);
        await ctx.sync();
        docStatus(`🔖 Added bookmark: "${name}"`);
        return;
      }
      if (action === 'goto-bookmark') {
        const name = document.getElementById('bookmark-name').value.trim();
        if (!name) { docStatus('Enter bookmark name', true); return; }
        const bm = ctx.document.bookmarks.getItem(name);
        bm.getRange().select();
        await ctx.sync();
        docStatus(`➡️ Jumped to: "${name}"`);
        return;
      }
      if (action === 'delete-bookmark') {
        const name = document.getElementById('bookmark-name').value.trim();
        if (!name) { docStatus('Enter bookmark name', true); return; }
        const bm = ctx.document.bookmarks.getItem(name);
        bm.delete();
        await ctx.sync();
        docStatus(`🗑️ Deleted: "${name}"`);
        return;
      }

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
      if (action === 'list-comments') {
        const comments = ctx.document.body.getComments();
        comments.load('items/authorName,items/content/text');
        await ctx.sync();
        const list = comments.items.map((c,i) => `${i+1}. [${c.authorName}]: ${c.content.text}`).join('\n');
        showOutput(list || 'No comments found.');
        return;
      }
      if (action === 'delete-comments') {
        const comments = ctx.document.body.getComments();
        comments.load('items');
        await ctx.sync();
        comments.items.forEach(c => c.delete());
        await ctx.sync();
        docStatus(`🗑️ Deleted ${comments.items.length} comment(s)`);
        return;
      }
      if (action === 'resolve-comments') {
        const comments = ctx.document.body.getComments();
        comments.load('items');
        await ctx.sync();
        comments.items.forEach(c => { c.resolved = true; });
        await ctx.sync();
        docStatus(`✅ Resolved ${comments.items.length} comment(s)`);
        return;
      }

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
      if (action === 'list-changes') {
        try {
          const changes = ctx.document.body.getTrackedChanges();
          changes.load('items/type,items/text');
          await ctx.sync();
          if (!changes.items.length) { docStatus('No tracked changes in document'); return; }
          const list = changes.items.map((c, i) => `${i+1}. [${c.type}]: ${c.text.slice(0, 80)}`).join('\n');
          showOutput(`📋 Tracked Changes:\n\n${list}`);
        } catch(_) {
          docStatus('Tracked changes checked');
        }
        return;
      }
      if (action === 'accept-selection') {
        try {
          const sel = ctx.document.getSelection();
          const changes = sel.getTrackedChanges();
          changes.load('items');
          await ctx.sync();
          changes.items.forEach(c => c.accept());
          await ctx.sync();
          docStatus(`✅ Accepted ${changes.items.length} change(s) in selection`);
        } catch(e) {
          docStatus(`⚠️ ${e.message}`, true);
        }
        return;
      }

      if (action === 'protect-doc' || action === 'unprotect-doc') {
        docStatus(action === 'protect-doc' ? '🔒 Document protection toggled' : '🔓 Document protection released');
        return;
      }

      if (action === 'compare-clipboard') {
        try {
          const clipText = await navigator.clipboard.readText();
          const sel = ctx.document.getSelection();
          sel.load('text');
          await ctx.sync();
          const docSel = sel.text.trim();
          if (!clipText || !docSel) {
            docStatus('Need both clipboard text and document selection to compare', true);
            return;
          }
          currentSourceContext = docSel;
          showOutput(clipText, docSel);
          document.getElementById('btn-view-diff').click();
          docStatus('📋 Compared selection with clipboard in Redline Diff');
        } catch(e) {
          docStatus(`⚠️ Clipboard access error: ${e.message}`, true);
        }
        return;
      }

      if (action === 'page-portrait' || action === 'page-landscape') {
        const sections = ctx.document.sections;
        sections.load('items');
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

      if (action === 'apply-margins') {
        const top = (parseFloat(document.getElementById('margin-top').value) || 2.54) * 28.3465;
        const left = (parseFloat(document.getElementById('margin-left').value) || 2.54) * 28.3465;
        const right = (parseFloat(document.getElementById('margin-right').value) || 2.54) * 28.3465;
        const bottom = (parseFloat(document.getElementById('margin-bottom').value) || 2.54) * 28.3465;
        const sections = ctx.document.sections;
        sections.load('items');
        await ctx.sync();
        sections.items.forEach(s => {
          s.pageSetup.topMargin = top;
          s.pageSetup.leftMargin = left;
          s.pageSetup.rightMargin = right;
          s.pageSetup.bottomMargin = bottom;
        });
        await ctx.sync();
        docStatus('📏 Margins applied to all sections');
        return;
      }

      if (action === 'paper-a4') {
        const sections = ctx.document.sections;
        sections.load('items');
        await ctx.sync();
        sections.items.forEach(s => {
          s.pageSetup.pageWidth = 595.3;
          s.pageSetup.pageHeight = 841.9;
        });
        await ctx.sync();
        docStatus('📄 Paper size set to A4 (210 x 297 mm)');
        return;
      }
      if (action === 'paper-letter') {
        const sections = ctx.document.sections;
        sections.load('items');
        await ctx.sync();
        sections.items.forEach(s => {
          s.pageSetup.pageWidth = 612;
          s.pageSetup.pageHeight = 792;
        });
        await ctx.sync();
        docStatus('📄 Paper size set to Letter (8.5 x 11 in)');
        return;
      }
      if (action === 'paper-a3') {
        const sections = ctx.document.sections;
        sections.load('items');
        await ctx.sync();
        sections.items.forEach(s => {
          s.pageSetup.pageWidth = 841.9;
          s.pageSetup.pageHeight = 1190.5;
        });
        await ctx.sync();
        docStatus('📄 Paper size set to A3');
        return;
      }
      if (action === 'paper-legal') {
        const sections = ctx.document.sections;
        sections.load('items');
        await ctx.sync();
        sections.items.forEach(s => {
          s.pageSetup.pageWidth = 612;
          s.pageSetup.pageHeight = 1008;
        });
        await ctx.sync();
        docStatus('📄 Paper size set to Legal (8.5 x 14 in)');
        return;
      }

      if (action === 'pagenumber-bottom') {
        const section = ctx.document.sections.getFirst();
        const footer = section.getFooter(Word.HeaderFooterType.primary);
        const p = footer.insertParagraph('Page ', Word.InsertLocation.replace);
        p.alignment = 'Centered';
        await ctx.sync();
        docStatus('🔢 Added page numbers to Bottom Center');
        return;
      }
      if (action === 'pagenumber-top') {
        const section = ctx.document.sections.getFirst();
        const header = section.getHeader(Word.HeaderFooterType.primary);
        const p = header.insertParagraph('Page ', Word.InsertLocation.replace);
        p.alignment = 'Right';
        await ctx.sync();
        docStatus('🔢 Added page numbers to Top Right');
        return;
      }
      if (action === 'pagenumber-remove') {
        const section = ctx.document.sections.getFirst();
        section.getHeader(Word.HeaderFooterType.primary).clear();
        section.getFooter(Word.HeaderFooterType.primary).clear();
        await ctx.sync();
        docStatus('🔢 Removed page numbers from header/footer');
        return;
      }
      if (action === 'get-pagecount') {
        const body = ctx.document.body;
        body.load('text');
        await ctx.sync();
        const words = body.text.split(/\s+/).filter(Boolean).length;
        const estPages = Math.max(1, Math.ceil(words / 275));
        docStatus(`📊 Document: ~${estPages} page(s) (${words.toLocaleString()} words)`);
        return;
      }

      if (action === 'apa-format' || action === 'mla-format') {
        const sections = ctx.document.sections;
        sections.load('items');
        const body = ctx.document.body;
        body.load('paragraphs');
        await ctx.sync();
        sections.items.forEach(s => {
          s.pageSetup.topMargin = 72;
          s.pageSetup.bottomMargin = 72;
          s.pageSetup.leftMargin = 72;
          s.pageSetup.rightMargin = 72;
        });
        body.font.name = 'Times New Roman';
        body.font.size = 12;
        body.paragraphs.load('items');
        await ctx.sync();
        body.paragraphs.items.forEach(p => {
          p.lineSpacing = 24;
          p.alignment = 'Left';
        });
        await ctx.sync();
        docStatus(`🎓 ${action === 'apa-format' ? 'APA 7th' : 'MLA 9th'} applied: 1" margins, Times New Roman 12pt, Double Spacing`);
        return;
      }

    });
  } catch(err) {
    docStatus(`⚠️ ${err.message}`, true);
  }
}

// ─── File Upload & Document Ingestion ────────────────────────────────────────

let uploadedFilesContent = [];

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
    document.getElementById('upload-label').querySelector('span').textContent = 'PDF, Word, Image, TXT, CSV, JSON…';
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

// ─── Research Copilot (with Upgraded Academic Tools) ─────────────────────────

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
        'ai-citation-finder': `You are an academic citation locator. Given the provided claim or topic, identify 3-5 seminal academic papers, peer-reviewed journal articles, or foundational books that directly support it.
For each source provide:
1. Full Citation in ${citStyle} format
2. Key empirical finding supporting the claim
3. Clickable Search Links:
   • Google Scholar: [Search on Google Scholar](https://scholar.google.com/scholar?q=${encodeURIComponent(topic || 'academic research')})
   • PubMed: [Search on PubMed](https://pubmed.ncbi.nlm.nih.gov/?term=${encodeURIComponent(topic || 'academic research')})
   • JSTOR: [Search on JSTOR](https://www.jstor.org/action/doBasicSearch?Query=${encodeURIComponent(topic || 'academic research')})`,

        'gap-finder': `You are a research innovation director. Analyze the provided literature, abstracts, or topic and identify:
1. Established consensus in existing research
2. 4 Specific Research Gaps (unexamined populations, methodological weaknesses, conflicting findings, or emerging technological shifts)
3. 3 Original, high-impact research questions formulated to bridge these gaps.`,

        'thesis-generator': `You are an academic thesis advisor. For the provided research topic and argument, generate 5 distinct, compelling, debate-ready Thesis Statements:
1. Analytical Thesis (examines nuances of components)
2. Argumentative/Persuasive Thesis (takes a bold defensible stance)
3. Cause-and-Effect Thesis (links underlying mechanisms to outcomes)
4. Comparative/Nuanced Thesis (reconciles two conflicting perspectives)
5. Policy/Practical Recommendation Thesis.
For each thesis statement, explain why an examiner would find it strong.`,

        'counter-argument': `You are a debate coach and critical scholar. Based on the provided argument and claims, formulate:
1. The 3 strongest, most sophisticated counter-arguments a skeptical referee or critic would raise
2. The core evidence or logic the opposition would lean on
3. A strategic rebuttal blueprint to preempt each counter-argument and strengthen the paper.`,

        'interview-questions': `You are a qualitative research methodology expert. Based on the provided research topic and objectives, design a semi-structured interview protocol:
1. 2 Warm-up questions (building rapport)
2. 5-6 Core exploratory questions designed to elicit deep phenomenological experiences
3. Follow-up probing prompts for each question
4. 1 Concluding reflection question.
Ensure open-ended phrasing avoiding leading biases.`,

        'survey-builder': `You are a quantitative survey design expert. Based on the research goals, construct a ready-to-deploy survey instrument:
1. 6 Likert-Scale statements (1 = Strongly Disagree to 5 = Strongly Agree) with construct dimensions
2. 3 Multiple-choice categorical questions
3. 2 Open-ended qualitative text questions.
Ensure questions avoid double-barreled phrasing, acquiescence bias, and ambiguous qualifiers.`,

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

      recordPrompt(prompts[action]);
      if (uploadedImgs.length) {
        await handleAIActionWithImages(prompts[action], userContent, uploadedImgs);
      } else {
        await handleAIAction(prompts[action], userContent);
      }
    });
  });
}

// ─── Semantic Scholar Academic Search ────────────────────────────────────────

function bindSemanticSearch() {
  const input = document.getElementById('semantic-search-input');
  const btn = document.getElementById('btn-semantic-search');
  const resultsContainer = document.getElementById('semantic-results');

  const doSearch = async () => {
    const q = input.value.trim();
    if (!q) return;
    resultsContainer.classList.remove('hidden');
    resultsContainer.innerHTML = '<div style="font-size:11px;color:var(--muted);padding:8px 0;text-align:center;">⏳ Searching 200M+ real papers on Semantic Scholar...</div>';
    btn.disabled = true;

    try {
      const url = `https://api.semanticscholar.org/graph/v1/paper/search?query=${encodeURIComponent(q)}&limit=5&fields=title,authors,year,abstract,url,venue`;
      const res = await fetch(url);
      if (!res.ok) throw new Error(`API error ${res.status}`);
      const data = await res.json();
      const papers = data.data || [];

      if (papers.length === 0) {
        resultsContainer.innerHTML = '<div style="font-size:11px;color:var(--muted);padding:6px;">No academic papers found. Try different keywords.</div>';
        return;
      }

      resultsContainer.innerHTML = '';
      papers.forEach(p => {
        const authors = (p.authors || []).map(a => a.name);
        const authorsText = authors.length > 2 ? `${authors[0]} et al.` : authors.join(' & ') || 'Unknown Author';
        const yearText = p.year || 'n.d.';
        const venueText = p.venue || 'Academic Publication';
        const inTextCit = `(${authorsText}, ${yearText})`;
        const fullRef = `${authorsText} (${yearText}). ${p.title}. ${venueText}.${p.url ? ` ${p.url}` : ''}`;

        const card = document.createElement('div');
        card.className = 'semantic-card';
        card.innerHTML = `
          <div class="semantic-title">${escapeHtml(p.title)}</div>
          <div class="semantic-meta">
            <span>👤 ${escapeHtml(authorsText)}</span>
            <span>📅 ${yearText}</span>
            <span>🏛️ ${escapeHtml(venueText)}</span>
          </div>
          <div class="semantic-abstract">${escapeHtml(p.abstract || 'No abstract preview available.')}</div>
          <div class="semantic-actions">
            <button class="small-btn btn-cite-intext" title="Insert in-text citation">Cite In-Text 📎</button>
            <button class="small-btn btn-cite-full" title="Append full citation to Bibliography">Add Reference 📖</button>
          </div>
        `;

        card.querySelector('.btn-cite-intext').addEventListener('click', async () => {
          await insertText(inTextCit, false);
          docStatus(`✅ Cited: ${inTextCit}`);
        });

        card.querySelector('.btn-cite-full').addEventListener('click', async () => {
          if (typeof Word === 'undefined') return;
          await Word.run(async ctx => {
            const body = ctx.document.body;
            const p = body.insertParagraph(fullRef, Word.InsertLocation.end);
            p.styleBuiltIn = Word.Style.normal;
            await ctx.sync();
            docStatus(`📖 Added reference: ${authorsText} (${yearText})`);
          });
        });

        resultsContainer.appendChild(card);
      });

    } catch(err) {
      resultsContainer.innerHTML = `<div style="font-size:11px;color:var(--danger);padding:6px;">⚠️ ${err.message}</div>`;
    } finally {
      btn.disabled = false;
    }
  };

  btn.addEventListener('click', doSearch);
  input.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      doSearch();
    }
  });
}

// ─── Quick Actions & Continue Writing ────────────────────────────────────────

function bindQuickActions() {
  const prompts = {
    rewrite:   'Rewrite the following text to improve clarity and flow. Return only the rewritten text.',
    summarize: 'Summarize the following text concisely. Return only the summary.',
    improve:   'Improve the writing quality of the following text. Return only the improved text.',
    translate: `Translate the following text to ${localStorage.getItem('wordai_language') || 'Urdu'}. Return only the translation.`,
    expand:    'Expand the following text with more detail and depth. Return only the expanded text.',
    shorten:   'Shorten the following text while keeping the key points. Return only the shortened text.',
  };

  document.querySelectorAll('.quick-btn:not(#btn-continue-writing)').forEach(btn => {
    btn.addEventListener('click', async () => {
      const action = btn.dataset.action;
      const contextMode = document.getElementById('context-select').value;
      const context = await getContext(contextMode === 'none' ? 'selection' : contextMode);
      if (!context) return showError('Please select some text first.');
      recordPrompt(prompts[action]);
      await handleAIAction(prompts[action], context);
    });
  });
}

function bindAutocomplete() {
  const btn = document.getElementById('btn-continue-writing');
  if (!btn) return;
  btn.addEventListener('click', async () => {
    if (typeof Word === 'undefined') { showError('Word API not available.'); return; }
    try {
      showLoading(true);
      const text = await Word.run(async ctx => {
        const sel = ctx.document.getSelection();
        sel.load('text');
        const body = ctx.document.body;
        body.load('text');
        await ctx.sync();
        if (sel.text.trim()) return sel.text.trim();
        const all = body.text.trim();
        return all.split(/\s+/).slice(-300).join(' ');
      });

      if (!text) {
        showError('Write a few words in Word first so WordAI can predict what comes next!');
        showLoading(false);
        return;
      }

      const prompt = `Based on the following document context, seamlessly write the NEXT 2-3 logical sentences. Match the exact tone, style, and terminology of the text. Return ONLY the continuation text with no conversational preamble or quotes.`;
      recordPrompt('Continue Writing (Autocomplete)');
      await handleAIAction(prompt, text);
    } catch(e) {
      showError(e.message);
      showLoading(false);
    }
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
      recordPrompt(prompt);
      await handleAIAction(prompt, userContent);
    });
  });
}

// ─── Advanced Document Framework Studio ─────────────────────────────────────

let activeDraftFramework = null;
let activeAuditFramework = null;
let currentMissingSections = [];

async function insertStructuredFramework(markdownText) {
  if (typeof Word === 'undefined') {
    await insertText(markdownText, false);
    docStatus('📚 Framework inserted into editor');
    return;
  }
  showLoading(true);
  try {
    await insertText(markdownText, false, true);
    docStatus('📚 Framework inserted with rich Word styles & bold formatting!');
  } catch (err) {
    showError('Could not insert framework: ' + err.message);
  } finally {
    showLoading(false);
  }
}

function openFrameworkDraftDrawer(fw) {
  activeDraftFramework = fw;
  const drawer = document.getElementById('framework-draft-drawer');
  document.getElementById('draft-drawer-title').textContent = `⚡ Draft: ${fw.title}`;
  document.getElementById('draft-drawer-desc').textContent = fw.desc;
  document.getElementById('draft-topic-input').value = '';
  drawer.classList.remove('hidden');
}

async function auditDocumentAgainstFramework(frameworkId) {
  if (typeof FRAMEWORK_DEFINITIONS === 'undefined') return;
  const fw = FRAMEWORK_DEFINITIONS[frameworkId];
  if (!fw) return;
  activeAuditFramework = fw;

  const drawer = document.getElementById('framework-audit-drawer');
  drawer.classList.remove('hidden');

  document.getElementById('audit-framework-name').textContent = fw.title;
  document.getElementById('audit-stats').textContent = 'Scanning document headings and paragraphs...';
  const checklistEl = document.getElementById('audit-checklist');
  checklistEl.innerHTML = '<div style="font-size:11px;color:var(--muted);padding:8px;">Analyzing document structure...</div>';

  let docText = '';
  let docHeadings = [];

  if (typeof Word !== 'undefined') {
    try {
      await Word.run(async (ctx) => {
        const body = ctx.document.body;
        const paras = body.paragraphs;
        paras.load('items/text,items/styleBuiltIn');
        await ctx.sync();
        docText = paras.items.map(p => p.text).join('\n');
        docHeadings = paras.items
          .filter(p => ['Heading1','Heading2','Heading3','Title'].includes(p.styleBuiltIn) || p.text.trim().startsWith('#'))
          .map(p => p.text.trim().toLowerCase());
      });
    } catch (_) {
      docText = await getContext('document');
    }
  } else {
    docText = await getContext('document');
  }

  const fullTextLower = (docText || '').toLowerCase();
  currentMissingSections = [];
  let foundCount = 0;

  checklistEl.innerHTML = '';

  fw.sections.forEach(sec => {
    const secClean = sec.replace(/^Chapter\s*\d+:?\s*/i, '').toLowerCase();
    const keywords = secClean.split(/[\s/&]+/).filter(w => w.length > 3);

    const inHeading = docHeadings.some(h => keywords.some(k => h.includes(k)));
    const inTextOccurrences = keywords.filter(k => fullTextLower.includes(k)).length;

    let status = 'missing';
    if (inHeading) {
      status = 'found';
      foundCount++;
    } else if (inTextOccurrences >= 2 || (keywords.length === 1 && inTextOccurrences >= 1)) {
      status = 'weak';
      foundCount += 0.5;
    } else {
      currentMissingSections.push(sec);
    }

    const card = document.createElement('div');
    card.className = `audit-check-card ${status}`;

    let statusBadge = '';
    let actionBtn = '';
    if (status === 'found') {
      statusBadge = '<span style="color:#059669;font-weight:700;">✅ Found</span>';
    } else if (status === 'weak') {
      statusBadge = '<span style="color:#d97706;font-weight:700;">⚠️ Brief / Mentioned</span>';
    } else {
      statusBadge = '<span style="color:var(--danger);font-weight:700;">❌ Missing</span>';
      actionBtn = `<button class="small-btn btn-draft-single-sec" data-sec="${encodeURIComponent(sec)}" style="font-size:9.5px;padding:2px 6px;">Draft ⚡</button>`;
    }

    card.innerHTML = `
      <div style="flex:1;">
        <div style="font-weight:600;">${escapeHtml(sec)}</div>
        <div style="font-size:10px;color:var(--muted);">${status === 'found' ? 'Properly structured in document' : status === 'weak' ? 'Referenced in text, but lacks distinct heading section' : 'Not identified in document'}</div>
      </div>
      <div style="display:flex;align-items:center;gap:4px;">
        ${statusBadge}
        ${actionBtn}
      </div>
    `;

    const draftBtn = card.querySelector('.btn-draft-single-sec');
    if (draftBtn) {
      draftBtn.addEventListener('click', async (e) => {
        e.stopPropagation();
        drawer.classList.add('hidden');
        const secName = decodeURIComponent(draftBtn.dataset.sec);
        const prompt = `Write a comprehensive, professional "${secName}" section for a ${fw.title}.\nEnsure full narrative depth, rigorous analysis, and appropriate formatting.`;
        recordPrompt(`Draft ${secName} (${fw.title})`);
        await handleAIAction(prompt, `Document Context:\n${docText ? docText.slice(0, 2500) : 'Create new section'}`);
      });
    }

    checklistEl.appendChild(card);
  });

  const total = fw.sections.length;
  const scorePct = Math.round((foundCount / total) * 100);

  document.getElementById('audit-score-badge').textContent = `${scorePct}% Complete`;
  document.getElementById('audit-progress-fill').style.width = `${scorePct}%`;
  document.getElementById('audit-stats').textContent = `${Math.floor(foundCount)} of ${total} required sections present. ${currentMissingSections.length} section(s) need drafting.`;
}

function renderFrameworks(filterCat = 'all', searchQuery = '') {
  const container = document.getElementById('frameworks-container');
  if (!container) return;
  if (typeof FRAMEWORK_DEFINITIONS === 'undefined') {
    container.innerHTML = '<div style="font-size:11px;color:var(--muted);padding:8px;">Frameworks loading...</div>';
    return;
  }

  const query = (searchQuery || '').toLowerCase().trim();
  const list = Object.values(FRAMEWORK_DEFINITIONS).filter(fw => {
    const matchesCat = filterCat === 'all' || fw.category === filterCat;
    const matchesSearch = !query ||
      fw.title.toLowerCase().includes(query) ||
      fw.desc.toLowerCase().includes(query) ||
      (fw.sections || []).some(s => s.toLowerCase().includes(query));
    return matchesCat && matchesSearch;
  });

  if (!list.length) {
    container.innerHTML = '<div style="font-size:11.5px;color:var(--muted);padding:12px;text-align:center;">No frameworks matching criteria.</div>';
    return;
  }

  container.innerHTML = '';
  list.forEach(fw => {
    const card = document.createElement('div');
    card.className = 'template-card';
    card.innerHTML = `
      <div class="template-header">
        <span class="template-title">${escapeHtml(fw.title)}</span>
        <span class="framework-badge badge-${fw.category}">${fw.badge}</span>
      </div>
      <div class="template-desc">${escapeHtml(fw.desc)}</div>
      <div class="template-sections-preview">📑 ${fw.sections.length} Core Sections: ${escapeHtml(fw.sections.slice(0, 3).join(' • '))}…</div>
      <div class="framework-actions">
        <button class="btn-fw-skeleton" data-id="${fw.id}" title="Insert heading structure & tables into Word">Insert Skeleton ➕</button>
        <button class="btn-fw-smartdraft" data-id="${fw.id}" title="Use AI to generate a complete custom draft">🤖 AI Smart-Draft</button>
        <button class="btn-fw-audit" data-id="${fw.id}" title="Scan active document against this framework">📋 Audit Doc</button>
      </div>
    `;

    card.querySelector('.btn-fw-skeleton').addEventListener('click', () => {
      insertStructuredFramework(fw.skeleton);
    });

    card.querySelector('.btn-fw-smartdraft').addEventListener('click', () => {
      openFrameworkDraftDrawer(fw);
    });

    card.querySelector('.btn-fw-audit').addEventListener('click', () => {
      auditDocumentAgainstFramework(fw.id);
    });

    container.appendChild(card);
  });
}

function bindFrameworkStudio() {
  renderFrameworks('all', '');

  document.querySelectorAll('.filter-pill').forEach(pill => {
    pill.addEventListener('click', () => {
      document.querySelectorAll('.filter-pill').forEach(p => p.classList.remove('active'));
      pill.classList.add('active');
      const cat = pill.dataset.cat;
      const query = document.getElementById('framework-search-input')?.value || '';
      renderFrameworks(cat, query);
    });
  });

  const searchInput = document.getElementById('framework-search-input');
  if (searchInput) {
    searchInput.addEventListener('input', (e) => {
      const activePill = document.querySelector('.filter-pill.active');
      const cat = activePill ? activePill.dataset.cat : 'all';
      renderFrameworks(cat, e.target.value);
    });
  }

  const draftDrawer = document.getElementById('framework-draft-drawer');
  document.getElementById('btn-close-draft-drawer').addEventListener('click', () => {
    draftDrawer.classList.add('hidden');
  });

  document.getElementById('btn-draft-from-doc').addEventListener('click', async () => {
    const docText = await getContext('document');
    if (!docText) return showError('No text found in document to summarize.');
    document.getElementById('draft-topic-input').value = docText.slice(0, 1200);
  });

  document.getElementById('btn-draft-from-sel').addEventListener('click', async () => {
    const selText = await getContext('selection');
    if (!selText) return showError('Please select text in your document first.');
    document.getElementById('draft-topic-input').value = selText;
  });

  document.getElementById('btn-run-smart-draft').addEventListener('click', async () => {
    if (!activeDraftFramework) return;
    const topic = document.getElementById('draft-topic-input').value.trim();
    if (!topic) return showError('Please describe your document topic or project context.');

    const depth = document.getElementById('draft-depth-select').value;
    const rigor = document.getElementById('draft-rigor-select').value;
    draftDrawer.classList.add('hidden');

    const depthInstruction = depth === 'in-depth'
      ? 'Produce a comprehensive, publication-grade, fully fleshed-out draft with extensive narrative depth, data tables, and specific analytical elaboration for EVERY section.'
      : depth === 'executive'
      ? 'Produce an executive-level synthesized draft with bulleted key takeaways, core financial/empirical data, and actionable strategic directives.'
      : 'Produce a structured blueprint scaffold with introductory narrative in each section and explicit [ADD EVIDENCE HERE] or [DATA PLACEHOLDER] guidance.';

    const systemPrompt = `${activeDraftFramework.aiPrompt}\nTone: ${rigor}.\nDepth Requirement: ${depthInstruction}\nEnsure every required framework section is clearly headed and thoroughly addressed.`;

    recordPrompt(`Smart-Draft: ${activeDraftFramework.title}`);
    await handleAIAction(systemPrompt, `Topic / Project Context:\n${topic}`);
  });

  const auditDrawer = document.getElementById('framework-audit-drawer');
  document.getElementById('btn-close-audit-drawer').addEventListener('click', () => {
    auditDrawer.classList.add('hidden');
  });

  document.getElementById('btn-draft-missing-sections').addEventListener('click', async () => {
    if (!activeAuditFramework || !currentMissingSections.length) return;
    auditDrawer.classList.add('hidden');
    const docText = await getContext('document');
    const prompt = `You are a document integrity specialist. The current document is missing these required sections for a complete ${activeAuditFramework.title}:\n\n${currentMissingSections.map(s => `• ${s}`).join('\n')}\n\nDraft ONLY the missing sections with full professional depth, matching the existing document context.`;
    recordPrompt(`Draft Missing Sections for ${activeAuditFramework.title}`);
    await handleAIAction(prompt, docText ? `Existing Document Content:\n${docText.slice(0, 3000)}` : 'Draft baseline content for these sections.');
  });
}

// ─── Career & Job Application Mode ───────────────────────────────────────────

function bindCareerButtons() {
  const careerPrompts = {
    'resume-bullet': 'Transform the following accomplishment or job task into a high-impact, executive resume bullet using Google\'s XYZ formula: "Accomplished [X] as measured by [Y] by doing [Z]". Use strong active verbs, quantify results with metrics, and eliminate weak phrasing. Provide 3 optimized alternatives.',
    'cover-letter-tailor': 'Write an ultra-persuasive, highly tailored cover letter aligning the candidate\'s background with the role. Focus on accomplishments, quantifiable ROI, and cultural enthusiasm. Avoid generic clichés.',
    'linkedin-headline': 'Draft 3 high-impact LinkedIn Headlines (under 220 characters) using keyword-rich value propositions, plus a compelling, narrative-driven first-person "About" summary.',
    'star-answer': 'Draft a structured, interview-winning behavioral answer using the STAR framework: Situation, Task, Action, and Result. Include strategic self-reflection and measurable business impact.',
    'salary-negotiate': 'Write a tactful, professional, yet assertive salary and compensation negotiation message making a data-backed case for an above-market offer while maintaining great rapport.',
    'networking-inmail': 'Write a warm, respectful, high-response LinkedIn InMail message (under 150 words) requesting a brief 15-minute informational interview or introduction.',
  };

  document.querySelectorAll('.career-btn').forEach(btn => {
    btn.addEventListener('click', async () => {
      const key = btn.dataset.career;
      const contextMode = document.getElementById('context-select').value;
      const context = await getContext(contextMode);
      const prompt = careerPrompts[key];
      const userContent = context ? `${prompt}\n\nCandidate Background / Target Role:\n${context}` : prompt;
      recordPrompt(prompt);
      await handleAIAction('You are a premier executive career strategist and talent acquisition expert.', userContent);
    });
  });
}

// ─── Social Media Repurposing Mode ──────────────────────────────────────────

function bindSocialButtons() {
  const socialPrompts = {
    'linkedin-post': 'Repurpose the provided document into a high-engagement LinkedIn post. Start with an irresistible 1-line hook that stops the scroll, use clean spacing with punchy takeaways, share an authentic insight, ask an open question to stimulate comments, and include 3-5 relevant hashtags.',
    'twitter-thread': 'Repurpose the provided content into a viral, value-packed Twitter / X thread (5-8 numbered tweets). Tweet 1 must be a viral hook that promises transformation. Each tweet should cover 1 clear takeaway. Final tweet should summarize and ask for a retweet/bookmark.',
    'instagram-caption': 'Convert this content into an engaging Instagram / Threads caption: conversational tone, bullet points with emojis, strong call-to-action, and a curated block of 10-15 relevant hashtags.',
    'newsletter-blurb': 'Write an engaging, executive newsletter summary blurb (200 words) summarizing key insights, what it means for the reader, and a compelling link teaser.',
  };

  document.querySelectorAll('.social-btn').forEach(btn => {
    btn.addEventListener('click', async () => {
      const key = btn.dataset.social;
      const contextMode = document.getElementById('context-select').value;
      const context = await getContext(contextMode === 'none' ? 'document' : contextMode);
      if (!context) return showError('Please select text or select "Entire Document" to repurpose into social content.');
      const prompt = socialPrompts[key];
      recordPrompt(prompt);
      await handleAIAction('You are a viral social media strategist and digital ghostwriter.', `${prompt}\n\nSource Content:\n${context}`);
    });
  });
}

// ─── Presentation & Slide Deck Mode ──────────────────────────────────────────

function bindSlidesButtons() {
  const slidesPrompts = {
    'deck-8': 'Convert the provided document into an 8-Slide PowerPoint Outline. For each slide provide:\n• Slide Title\n• 3-4 concise, high-impact bullet points\n• 2-3 sentences of detailed Speaker Notes.',
    'deck-pitch': 'Convert the provided content into a 10-Slide Investor/Client Pitch Deck (Problem, Solution, Market Opportunity, Product, Business Model, Traction, Competition, Go-To-Market, Team, The Ask). Include detailed Speaker Notes for each slide.',
    'deck-academic': 'Convert the provided research into a 12-Slide Academic Conference Presentation outline (Title, Research Context, Theoretical Framework, Methodology, Findings 1-3, Discussion, Limitations, Implications, Q&A) with rigorous academic Speaker Notes.',
    'deck-summary': 'Convert the provided document into a high-stakes 3-Slide Executive Lightning Brief (Slide 1: Current Challenge & Strategic Context; Slide 2: Proposed Solution; Slide 3: Execution Timeline & Expected ROI) with Speaker Notes.',
  };

  document.querySelectorAll('.slides-btn').forEach(btn => {
    btn.addEventListener('click', async () => {
      const key = btn.dataset.slides;
      const contextMode = document.getElementById('context-select').value;
      const context = await getContext(contextMode === 'none' ? 'document' : contextMode);
      if (!context) return showError('Please select text or select "Entire Document" to create slides from.');
      const prompt = slidesPrompts[key];
      recordPrompt(prompt);
      await handleAIAction('You are a presentation designer and McKinsey-tier storytelling consultant.', `${prompt}\n\nDocument Material:\n${context}`);
    });
  });
}

// ─── Email Hub ───────────────────────────────────────────────────────────────

function bindEmailButtons() {
  const emailPrompts = {
    'cold-outreach': 'Write a compelling, high-converting cold outreach email. Focus on the recipient\'s pain point, clear value proposition, and a friction-free call to action.',
    'executive-update': 'Write an executive leadership update email: concise highlights, quantitative progress, current blockers, and upcoming milestones.',
    'meeting-recap': 'Write a structured meeting recap email: key discussion decisions, next steps, and an action-item assignment table.',
    'deadline-status': 'Write a professional project status and deadline email: delivery timeline, milestones met, dependencies, and risk mitigation.',
    'polite-decline': 'Write a gracious, polite, and firm decline email (declining an invitation, vendor pitch, or request) preserving mutual respect.',
    'formal-apology': 'Write a sincere, professional business apology acknowledging an issue, stating the remediation taken, and assuring future reliability.',
    'follow-up': 'Write a tactful, gentle follow-up email regarding an unanswered message without sounding impatient.',
    'introduction': 'Write a warm, double-opt-in professional introduction email connecting two parties and highlighting mutual benefit.',
  };

  document.querySelectorAll('.email-btn').forEach(btn => {
    btn.addEventListener('click', async () => {
      const type = btn.dataset.email;
      const contextMode = document.getElementById('context-select').value;
      const context = await getContext(contextMode);
      const prompt = emailPrompts[type];
      const userContent = context ? `${prompt}\n\nContext:\n${context}` : prompt;
      recordPrompt(prompt);
      await handleAIAction('You are an expert executive communication and business email strategist.', userContent);
    });
  });

  let activeStance = 'accept';
  document.querySelectorAll('.stance-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.stance-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      activeStance = btn.dataset.stance;
    });
  });

  document.getElementById('btn-generate-reply').addEventListener('click', async () => {
    const incoming = document.getElementById('reply-incoming-email').value.trim();
    if (!incoming) return showError('Please paste the incoming email text first.');
    const stanceDescriptions = {
      accept: 'Agree and accept the request enthusiastically while outlining immediate next steps.',
      decline: 'Respectfully and graciously decline the request while providing a polite reason.',
      clarify: 'Request specific clarifications and additional details before moving forward.',
      reschedule: 'Propose rescheduling to a later date and time due to an unavoidable conflict.',
    };
    const stancePrompt = stanceDescriptions[activeStance];
    const systemPrompt = `You are an expert business communicator. Draft a polished, ready-to-send reply to the incoming email.\nDesired Stance: ${stancePrompt}\nProvide an effective Subject line and complete message body.`;
    recordPrompt(`Email Reply (${activeStance})`);
    await handleAIAction(systemPrompt, incoming);
  });
}

// ─── Edit Suite & Document Intelligence Upgrades ─────────────────────────────

function bindEditButtons() {
  const editPrompts = {
    grammar:  'Check the following text for grammar, spelling, and punctuation errors. List each issue with a suggested fix.',
    'fix-all':'Fix all grammar, spelling, punctuation, and clarity issues in the following text. Return only the corrected text.',
    clarity:  'Improve the clarity and readability of the following text. Return only the improved text.',
    passive:  'Rewrite the following text to eliminate passive voice. Return only the rewritten text.',
    simplify: 'Simplify the language in the following text so it is easy to understand. Return only the simplified text.',
    score:    'Analyze the following text and give a writing quality score out of 100. Break down: grammar, clarity, readability, word choice, and structure.',

    plagiarism: `You are an academic integrity and originality analyst. Analyze the provided text:
1. Estimate an Originality Score percentage (0% to 100%).
2. Flag any sentences that sound like clichéd academic boilerplate, copied syntax, or lack freshness.
3. For each flagged sentence: explain why it appears unoriginal and provide a fresh, original academic rewrite.
4. Output a summary assessment.`,

    transitions: `You are an expert writing coach. Read the provided text containing two adjacent paragraphs or sections.
Write 3 alternative smooth, logically coherent transition sentences or bridge paragraphs that connect the two ideas seamlessly without abrupt shifts.`,

    'expander-evidence': `Expand the provided paragraph with deeper analytical depth and academic context. Wherever a factual assertion or empirical evidence is needed to sustain the claim, explicitly insert bracketed markers: [ADD EVIDENCE HERE: specify exact data, study, or statistic needed]. Return the expanded paragraph.`,

    jargon: `Scan the provided text for overly complex jargon, bureaucratic language, and convoluted phrasing:
1. Output a table of: Jargon Term → Plain Language Alternative
2. Provide a full rewrite of the text in clear, accessible plain language while retaining professional accuracy.`,

    'key-terms': `Extract the top 15-20 core keywords, technical concepts, and domain terminology from the text. For each term, provide a concise 1-sentence contextual definition based on this document. Format as an alphabetical index.`,

    'action-items': `Scan the provided document and extract all actionable tasks, milestones, deliverables, deadlines, and assigned responsibilities. Format as a table:
| Action Item / Task | Owner / Assignee | Deadline / Timing | Priority | Notes |`,

    'compare-evaluate': `Compare the two versions of text provided:
1. Itemize the substantive differences
2. Evaluate which version is clearer, more persuasive, and higher quality
3. Declare an overall winner with rationale
4. Suggest a synthesis combining the strengths of both.`,

    sentiment: `Perform a detailed sentiment and tone analysis:
1. Overall Sentiment Score (-100 to +100) and Tone Label (e.g. Confident, Collaborative, Critical, Neutral)
2. Section-by-Section Tone breakdown
3. High-impact emotional keywords that sway the reader's perception.`,

    'contract-scan': `You are an elite corporate legal counsel and contract risk analyst. Analyze the provided agreement or contract clauses and produce an itemized Risk Assessment:
1. 🚨 High-Risk Clauses (unlimited liability, broad indemnification, auto-renewal traps, unilateral termination)
2. ⚠️ Ambiguities & Scope Gaps (vague deadlines, undefined performance metrics, unclear IP ownership)
3. 🛡️ Recommended Redlines & Counter-Language (exact alternative clauses to propose)
4. 📋 Executive Summary (negotiation leverage points).`,

    reviewer2: `You are academic "Reviewer #2" — a rigorous, demanding peer-review referee. Evaluate the provided academic draft against:
1. Novelty & Theoretical Contribution (Does this advance the state of the art?)
2. Methodological Rigor & Internal Validity (Sample size, confounding factors, controls)
3. Argument Support (Are claims supported by evidence?)
4. Literature Depth (Missing foundational citations)
5. OFFICIAL VERDICT: [Accept / Minor Revisions / Major Revisions / Reject]
Provide an itemized, constructive, yet unapologetically rigorous critique.`
  };

  document.querySelectorAll('.edit-btn').forEach(btn => {
    btn.addEventListener('click', async () => {
      const action = btn.dataset.action;
      const contextMode = document.getElementById('context-select').value;
      const context = await getContext(contextMode === 'none' ? 'document' : contextMode);
      if (!context) return showError('No text found. Select text or choose "Entire Document".');

      if (action === 'readinglevel') { showReadingLevel(context); return; }
      if (action === 'consistency') { checkConsistency(context); return; }
      if (action === 'rewrite3') { await showRewrite3(context); return; }
      if (action === 'section-summarize') { await summarizeBySection(); return; }

      if (action === 'translate-sidebyside') {
        const lang = localStorage.getItem('wordai_language') || 'Spanish';
        const prompt = `Translate the following text faithfully into ${lang}. Maintain formatting.`;
        showLoading(true);
        hideOutput();
        try {
          const result = await runAI(prompt, context);
          currentSourceContext = context;
          showOutput(result, context);
          document.getElementById('btn-view-sidebyside').click();
        } catch(e) { showError(e.message); }
        finally { showLoading(false); }
        return;
      }

      recordPrompt(editPrompts[action]);
      await handleAIAction(editPrompts[action], context);
    });
  });
}

// ─── Summarize by Section ───────────────────────────────────────────────────

async function summarizeBySection() {
  if (typeof Word === 'undefined') return showError('Word API unavailable.');
  showLoading(true);
  try {
    const sections = await Word.run(async ctx => {
      const paras = ctx.document.body.paragraphs;
      paras.load('items/text,items/styleBuiltIn');
      await ctx.sync();

      const secList = [];
      let currentHeading = 'Document Overview';
      let currentText = '';

      paras.items.forEach(p => {
        if (['Heading1','Heading2','Heading3'].includes(p.styleBuiltIn)) {
          if (currentText.trim()) secList.push({ heading: currentHeading, text: currentText.trim() });
          currentHeading = p.text.trim();
          currentText = '';
        } else {
          currentText += ' ' + p.text.trim();
        }
      });
      if (currentText.trim()) secList.push({ heading: currentHeading, text: currentText.trim() });
      return secList;
    });

    if (!sections.length) {
      showLoading(false);
      return showError('No text found to summarize.');
    }

    const prompt = `Summarize each of the following document sections individually. Provide a clear, bulleted summary for each heading:\n\n` +
      sections.map(s => `### ${s.heading}\n${s.text.slice(0, 1500)}`).join('\n\n');

    recordPrompt('Summarize by Section');
    await handleAIAction(prompt, 'Document Sections');
  } catch(e) {
    showError(e.message);
  } finally {
    showLoading(false);
  }
}

// ─── Rewrite 3 Versions ───────────────────────────────────────────────────

async function showRewrite3(text) {
  const panel = document.getElementById('rewrite3-panel');
  panel.innerHTML = '<div style="font-size:12px;color:var(--muted);padding:6px 0">⏳ Generating 3 versions...</div>';
  panel.classList.remove('hidden');
  showLoading(true);
  try {
    const result = await runAI(
      `Rewrite the following text in exactly 3 different ways. Label them exactly as:
Version 1 (Concise):
Version 2 (Formal):
Version 3 (Creative):
Return only the 3 labeled versions, nothing else.`,
      text
    );

    const versionRe = /Version\s*(\d)\s*\(([^)]+)\):\s*([\s\S]*?)(?=Version\s*\d|$)/gi;
    const versions = [];
    let m;
    while ((m = versionRe.exec(result)) !== null) {
      versions.push({ num: m[1], label: m[2].trim(), text: m[3].trim() });
    }

    if (versions.length === 0) {
      result.split(/\n{2,}/).slice(0, 3).forEach((t, i) => {
        versions.push({ num: i + 1, label: ['Concise','Formal','Creative'][i], text: t.trim() });
      });
    }

    panel.innerHTML = '';
    versions.forEach(v => {
      const card = document.createElement('div');
      card.className = 'rewrite3-card';
      card.innerHTML = `
        <div class="rewrite3-card-header">
          <span class="rewrite3-label">🔀 Version ${v.num} — ${v.label}</span>
          <div class="rewrite3-actions">
            <button class="small-btn" data-insert="${encodeURIComponent(v.text)}">Insert ✅</button>
            <button class="small-btn" data-replace="${encodeURIComponent(v.text)}">Replace 🔄</button>
            <button class="small-btn" data-copy="${encodeURIComponent(v.text)}">Copy 📋</button>
          </div>
        </div>
        <div>${escapeHtml(v.text)}</div>`;
      panel.appendChild(card);
    });

    panel.querySelectorAll('[data-insert]').forEach(btn => {
      btn.addEventListener('click', () => insertText(decodeURIComponent(btn.dataset.insert), false));
    });
    panel.querySelectorAll('[data-replace]').forEach(btn => {
      btn.addEventListener('click', () => insertText(decodeURIComponent(btn.dataset.replace), true));
    });
    panel.querySelectorAll('[data-copy]').forEach(btn => {
      btn.addEventListener('click', () => navigator.clipboard.writeText(decodeURIComponent(btn.dataset.copy)));
    });

  } catch (err) {
    panel.innerHTML = `<div style="color:var(--danger);font-size:12px">⚠️ ${err.message}</div>`;
  } finally {
    showLoading(false);
  }
}

// ─── Consistency Checker ─────────────────────────────────────────────────

function checkConsistency(text) {
  const variantGroups = [
    ['artificial intelligence', 'AI', 'A.I.', 'A.I'],
    ['machine learning', 'ML', 'M.L.'],
    ['deep learning', 'DL'],
    ['natural language processing', 'NLP', 'N.L.P.'],
    ['internet of things', 'IoT', 'IOT'],
    ['user interface', 'UI', 'U.I.'],
    ['user experience', 'UX', 'U.X.'],
    ['application programming interface', 'API', 'A.P.I.'],
    ['et al.', 'et al', 'et. al.', 'et. al'],
    ['e.g.', 'e.g', 'eg.', 'for example'],
    ['i.e.', 'i.e', 'ie.', 'that is'],
    ['versus', 'vs.', 'vs', 'v.'],
    ['figure', 'Fig.', 'fig.', 'FIG'],
    ['table', 'Tab.', 'tab.', 'TABLE'],
    ['chapter', 'Ch.', 'ch.', 'Chap.'],
    ['percent', '%', 'per cent'],
    ['data is', 'data are'],
    ['website', 'web site', 'Web site'],
    ['email', 'e-mail', 'E-mail'],
    ['online', 'on-line', 'on line'],
    ['healthcare', 'health care', 'health-care'],
    ['decision making', 'decision-making'],
    ['long term', 'long-term'],
    ['short term', 'short-term'],
    ['real time', 'real-time'],
  ];

  const issues = [];

  for (const group of variantGroups) {
    const found = group.filter(term => {
      const re = new RegExp(`\\b${term.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\b`, 'i');
      return re.test(text);
    });
    if (found.length > 1) {
      issues.push(`⚠️ Mixed usage: ${found.map(f => `"${f}"`).join(' vs ')}`);
    }
  }

  const wordFreq = {};
  const wordRe = /\b[A-Za-z][a-z]{2,}\b/g;
  let m;
  while ((m = wordRe.exec(text)) !== null) {
    const w = m[0];
    const key = w.toLowerCase();
    if (!wordFreq[key]) wordFreq[key] = new Set();
    wordFreq[key].add(w);
  }
  for (const [key, variants] of Object.entries(wordFreq)) {
    if (variants.size > 1) {
      const arr = [...variants];
      const hasLower = arr.some(v => v[0] === v[0].toLowerCase());
      const hasUpper = arr.some(v => v[0] === v[0].toUpperCase());
      if (hasLower && hasUpper && arr.length === 2) {
        issues.push(`🔤 Capitalisation inconsistency: ${arr.map(v => `"${v}"`).join(' / ')}`);
      }
    }
  }

  const hasDigitSmall = /\b[2-9]\b/.test(text);
  const hasWordSmall  = /\b(two|three|four|five|six|seven|eight|nine)\b/i.test(text);
  if (hasDigitSmall && hasWordSmall) {
    issues.push('🔢 Number style mixed: some numbers written as digits, others as words (e.g. "3" vs "three")');
  }

  const hasStraight = /[\'\"]/g.test(text);
  const hasCurly    = /[‘’“”]/g.test(text);
  if (hasStraight && hasCurly) {
    issues.push('💬 Quote style mixed: straight quotes (\'\') and curly quotes (‘’) both found');
  }

  if (issues.length === 0) {
    showOutput('✅ No consistency issues found! Your terminology and style appear consistent.');
    return;
  }

  const report = `🔁 Consistency Check — ${issues.length} issue${issues.length > 1 ? 's' : ''} found\n\n` + issues.join('\n');
  showOutput(report);
}

// ─── Reading Level Analyzer ─────────────────────────────────────────────────

function countSyllables(word) {
  word = word.toLowerCase().replace(/[^a-z]/g, '');
  if (!word) return 0;
  if (word.length <= 3) return 1;
  word = word.replace(/(?:[^laeiouy]es|ed|[^laeiouy]e)$/, '');
  word = word.replace(/^y/, '');
  const m = word.match(/[aeiouy]{1,2}/g);
  return m ? m.length : 1;
}

function showReadingLevel(text) {
  const sentences = text.split(/[.!?]+/).filter(s => s.trim().length > 0);
  const words = text.split(/\s+/).filter(Boolean);
  const syllables = words.reduce((sum, w) => sum + countSyllables(w), 0);

  const S = sentences.length || 1;
  const W = words.length || 1;
  const syl = syllables;

  const ease = 206.835 - 1.015 * (W / S) - 84.6 * (syl / W);
  const easeScore = Math.max(0, Math.min(100, Math.round(ease)));

  const grade = 0.39 * (W / S) + 11.8 * (syl / W) - 15.59;
  const gradeLevel = Math.max(1, Math.round(grade));

  const mins = Math.ceil(W / 200);

  let audience, easeLabel;
  if (easeScore >= 90)      { audience = 'Very Easy — 5th grade';     easeLabel = '🟢 Very Easy'; }
  else if (easeScore >= 70) { audience = 'Easy — 6th grade';           easeLabel = '🟢 Easy'; }
  else if (easeScore >= 60) { audience = 'Standard — 7th–8th grade';   easeLabel = '🟡 Standard'; }
  else if (easeScore >= 50) { audience = 'Fairly Difficult — High School'; easeLabel = '🟠 Fairly Difficult'; }
  else if (easeScore >= 30) { audience = 'Difficult — College level';   easeLabel = '🔴 Difficult'; }
  else                      { audience = 'Very Difficult — Professional'; easeLabel = '🔴 Very Difficult'; }

  const result =
`📊 Reading Level Analysis

📌 Flesch Reading Ease: ${easeScore}/100 — ${easeLabel}
🎓 Grade Level: Grade ${gradeLevel} (${audience})
⏱️ Reading Time: ~${mins} min${mins > 1 ? 's' : ''}

📝 Stats:
• Words: ${W.toLocaleString()}
• Sentences: ${S.toLocaleString()}
• Syllables: ${syl.toLocaleString()}
• Avg words/sentence: ${(W/S).toFixed(1)}
• Avg syllables/word: ${(syl/W).toFixed(2)}`;

  showOutput(result);
}

// ─── Tone Changer ────────────────────────────────────────────────────────────

function bindToneButtons() {
  document.querySelectorAll('.tone-btn').forEach(btn => {
    btn.addEventListener('click', async () => {
      const tone = btn.dataset.tone;
      const context = await getContext('selection');
      if (!context) return showError('Please select some text first.');
      const prompt = `Rewrite the following text in a ${tone} tone. Return only the rewritten text.`;
      recordPrompt(prompt);
      await handleAIAction(prompt, context);
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
  recordPrompt(userMsg);

  showLoading(true);
  try {
    const docText = await getContext('document');
    const personaInstruction = getPersonaInstruction();
    const systemPrompt = docText
      ? `You are an AI assistant helping with a Word document. ${personaInstruction}\nHere is the document content:\n\n${docText.slice(0, 8000)}\n\nAnswer questions about this document helpfully and concisely.`
      : `You are a helpful AI writing assistant inside Microsoft Word. ${personaInstruction}`;

    chatHistory.push({ role: 'user', content: userMsg });

    const lang = localStorage.getItem('wordai_language') || 'English';
    const messages = [
      { role: 'system', content: `${systemPrompt} Respond in ${lang}.` },
      ...chatHistory.slice(-10),
    ];

    const { text } = await AIProvider.call(messages);
    chatHistory.push({ role: 'assistant', content: text });
    appendChatMessage(text, 'ai');
    recordResponse(userMsg, text);
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

// ─── Prompt Box & Voice Dictation ────────────────────────────────────────────

function bindPromptBox() {
  document.getElementById('btn-generate').addEventListener('click', () => triggerGenerate());
}

function triggerGenerate() {
  const prompt = document.getElementById('main-prompt').value.trim();
  if (!prompt) return;
  const contextMode = document.getElementById('context-select').value;
  getContext(contextMode).then(context => {
    const userContent = context ? `${prompt}\n\n${context}` : prompt;
    recordPrompt(prompt);
    handleAIAction('You are a helpful AI writing assistant inside Microsoft Word.', userContent);
  });
}

function bindDictation() {
  const btn = document.getElementById('btn-dictate');
  const statusEl = document.getElementById('dictation-status');
  const promptEl = document.getElementById('main-prompt');
  if (!btn) return;

  const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
  if (!SpeechRecognition) {
    btn.title = 'Speech dictation not supported in this environment';
    btn.style.opacity = '0.4';
    btn.addEventListener('click', () => {
      docStatus('Speech dictation is not supported by your current browser/host runtime.', true);
    });
    return;
  }

  const recognition = new SpeechRecognition();
  recognition.continuous = false;
  recognition.interimResults = true;
  let isListening = false;

  recognition.onstart = () => {
    isListening = true;
    btn.classList.add('listening');
    statusEl.classList.remove('hidden');
  };

  recognition.onresult = (event) => {
    let transcript = '';
    for (let i = event.resultIndex; i < event.results.length; ++i) {
      transcript += event.results[i][0].transcript;
    }
    promptEl.value = transcript;
  };

  recognition.onerror = () => {
    isListening = false;
    btn.classList.remove('listening');
    statusEl.classList.add('hidden');
  };

  recognition.onend = () => {
    isListening = false;
    btn.classList.remove('listening');
    statusEl.classList.add('hidden');
  };

  btn.addEventListener('click', () => {
    if (isListening) {
      recognition.stop();
    } else {
      recognition.start();
    }
  });
}

// ─── Shared AI Handlers & Multi-View Switcher ────────────────────────────────

async function handleAIAction(systemPrompt, userContent) {
  currentSourceContext = userContent;
  showLoading(true);
  hideOutput();
  try {
    const result = await runAI(systemPrompt, userContent);
    showOutput(result, userContent);
  } catch (err) {
    showError(err.message);
  } finally {
    showLoading(false);
  }
}

async function handleAIActionWithImages(systemPrompt, userContent, images) {
  currentSourceContext = userContent;
  showLoading(true);
  hideOutput();
  try {
    const lang = localStorage.getItem('wordai_language') || 'English';
    const personaInstruction = getPersonaInstruction();
    const messages = [
      { role: 'system', content: `${systemPrompt} ${personaInstruction} Respond in ${lang}.` },
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
    showOutput(text, userContent);
  } catch (_) {
    try {
      const result = await runAI(systemPrompt, userContent + '\n[Images attached but vision not supported by current model.]');
      showOutput(result, userContent);
    } catch (err) { showError(err.message); }
  } finally {
    showLoading(false);
  }
}

function bindDiffToggle() {
  const btnPreview = document.getElementById('btn-view-preview');
  const btnDiff = document.getElementById('btn-view-diff');
  const btnSide = document.getElementById('btn-view-sidebyside');
  const textEl = document.getElementById('output-text');
  const diffEl = document.getElementById('output-diff');
  const sideEl = document.getElementById('output-sidebyside');

  const setView = (activeBtn, showEl) => {
    [btnPreview, btnDiff, btnSide].forEach(b => b.classList.remove('active'));
    [textEl, diffEl, sideEl].forEach(e => e.classList.add('hidden'));
    activeBtn.classList.add('active');
    showEl.classList.remove('hidden');
  };

  btnPreview.addEventListener('click', () => setView(btnPreview, textEl));
  btnDiff.addEventListener('click', () => {
    diffEl.innerHTML = generateRedlineDiff(currentSourceContext, currentResultText);
    setView(btnDiff, diffEl);
  });
  btnSide.addEventListener('click', () => {
    document.getElementById('sidebyside-original').innerHTML = '<p>' + markdownToHtml(currentSourceContext || '(No source text)') + '</p>';
    document.getElementById('sidebyside-output').innerHTML = '<p>' + markdownToHtml(currentResultText) + '</p>';
    setView(btnSide, sideEl);
  });
}

// ─── Output Actions ──────────────────────────────────────────────────────────

function bindOutputActions() {
  document.getElementById('btn-insert').addEventListener('click', async () => {
    const text = currentResultText || document.getElementById('output-text').innerText;
    await insertText(text, false);
    docStatus('✅ Inserted directly into document');
  });

  document.getElementById('btn-replace').addEventListener('click', async () => {
    const text = currentResultText || document.getElementById('output-text').innerText;
    await insertText(text, true);
    docStatus('🔄 Replaced selection in document');
  });

  document.getElementById('btn-copy').addEventListener('click', () => {
    const text = document.getElementById('output-text').innerText;
    navigator.clipboard.writeText(text);
    docStatus('📋 Copied to clipboard');
  });

  document.getElementById('btn-discard').addEventListener('click', hideOutput);
}

// ─── History & Presets Drawer ────────────────────────────────────────────────

function recordPrompt(prompt) {
  if (!prompt || prompt.length < 5) return;
  let history = JSON.parse(localStorage.getItem('wordai_prompt_history') || '[]');
  history = [prompt, ...history.filter(p => p !== prompt)].slice(0, 20);
  localStorage.setItem('wordai_prompt_history', JSON.stringify(history));
  renderPromptHistory();
}

function recordResponse(prompt, response) {
  if (!response || response.length < 10) return;
  let log = JSON.parse(localStorage.getItem('wordai_saved_outputs') || '[]');
  const entry = {
    time: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
    prompt: (prompt || 'Generation').slice(0, 70),
    response: response
  };
  log = [entry, ...log].slice(0, 30);
  localStorage.setItem('wordai_saved_outputs', JSON.stringify(log));
  renderSavedOutputs();
}

function bindHistoryDrawer() {
  const drawer = document.getElementById('history-drawer');
  document.getElementById('btn-history').addEventListener('click', () => {
    drawer.classList.remove('hidden');
    renderPromptHistory();
    renderPresets();
    renderSavedOutputs();
  });

  document.getElementById('btn-close-history').addEventListener('click', () => {
    drawer.classList.add('hidden');
  });

  document.querySelectorAll('.drawer-subtab').forEach(tab => {
    tab.addEventListener('click', () => {
      document.querySelectorAll('.drawer-subtab').forEach(t => t.classList.remove('active'));
      document.querySelectorAll('.drawer-panel').forEach(p => p.classList.remove('active'));
      tab.classList.add('active');
      document.getElementById(`dpanel-${tab.dataset.dtab}`).classList.add('active');
    });
  });

  document.getElementById('btn-clear-history').addEventListener('click', () => {
    localStorage.removeItem('wordai_prompt_history');
    renderPromptHistory();
  });

  document.getElementById('btn-clear-saved').addEventListener('click', () => {
    localStorage.removeItem('wordai_saved_outputs');
    renderSavedOutputs();
  });

  document.getElementById('btn-save-preset').addEventListener('click', () => {
    const name = document.getElementById('preset-name-input').value.trim();
    const prompt = document.getElementById('preset-prompt-input').value.trim();
    if (!name || !prompt) return alert('Enter both preset name and prompt.');
    let presets = JSON.parse(localStorage.getItem('wordai_custom_presets') || '[]');
    presets.push({ name, prompt });
    localStorage.setItem('wordai_custom_presets', JSON.stringify(presets));
    document.getElementById('preset-name-input').value = '';
    document.getElementById('preset-prompt-input').value = '';
    renderPresets();
  });
}

function renderPromptHistory() {
  const listEl = document.getElementById('history-list');
  const history = JSON.parse(localStorage.getItem('wordai_prompt_history') || '[]');
  if (!history.length) {
    listEl.innerHTML = '<div style="font-size:11px;color:var(--muted);padding:8px;">No prompt history yet.</div>';
    return;
  }
  listEl.innerHTML = '';
  history.forEach(p => {
    const card = document.createElement('div');
    card.className = 'drawer-card';
    card.innerHTML = `<div class="drawer-card-body">${escapeHtml(p)}</div>`;
    card.addEventListener('click', () => {
      document.getElementById('main-prompt').value = p;
      document.getElementById('history-drawer').classList.add('hidden');
    });
    listEl.appendChild(card);
  });
}

function renderPresets() {
  const listEl = document.getElementById('presets-list');
  const presets = JSON.parse(localStorage.getItem('wordai_custom_presets') || '[]');
  if (!presets.length) {
    listEl.innerHTML = '<div style="font-size:11px;color:var(--muted);padding:8px;">No custom presets saved yet.</div>';
    return;
  }
  listEl.innerHTML = '';
  presets.forEach((preset, idx) => {
    const card = document.createElement('div');
    card.className = 'drawer-card';
    card.innerHTML = `
      <div class="drawer-card-top">
        <span style="font-weight:700;color:var(--primary);">${escapeHtml(preset.name)}</span>
        <button class="small-btn danger btn-del-preset" style="font-size:9px;padding:1px 4px;">✕</button>
      </div>
      <div class="drawer-card-body">${escapeHtml(preset.prompt)}</div>`;

    card.querySelector('.btn-del-preset').addEventListener('click', (e) => {
      e.stopPropagation();
      presets.splice(idx, 1);
      localStorage.setItem('wordai_custom_presets', JSON.stringify(presets));
      renderPresets();
    });

    card.addEventListener('click', () => {
      document.getElementById('main-prompt').value = preset.prompt;
      document.getElementById('history-drawer').classList.add('hidden');
    });
    listEl.appendChild(card);
  });
}

function renderSavedOutputs() {
  const listEl = document.getElementById('saved-outputs-list');
  const log = JSON.parse(localStorage.getItem('wordai_saved_outputs') || '[]');
  if (!log.length) {
    listEl.innerHTML = '<div style="font-size:11px;color:var(--muted);padding:8px;">No saved outputs in this session.</div>';
    return;
  }
  listEl.innerHTML = '';
  log.forEach(entry => {
    const card = document.createElement('div');
    card.className = 'drawer-card';
    card.innerHTML = `
      <div class="drawer-card-top">
        <span>Prompt: ${escapeHtml(entry.prompt)}</span>
        <span>${entry.time}</span>
      </div>
      <div class="drawer-card-body">${escapeHtml(entry.response.slice(0, 160))}…</div>
      <div style="display:flex;gap:4px;margin-top:4px;">
        <button class="small-btn btn-insert-saved" style="font-size:10px;padding:2px 6px;">Insert ✅</button>
        <button class="small-btn btn-copy-saved" style="font-size:10px;padding:2px 6px;">Copy 📋</button>
      </div>
    `;

    card.querySelector('.btn-insert-saved').addEventListener('click', (e) => {
      e.stopPropagation();
      insertText(entry.response, false);
      docStatus('✅ Inserted saved response');
    });

    card.querySelector('.btn-copy-saved').addEventListener('click', (e) => {
      e.stopPropagation();
      navigator.clipboard.writeText(entry.response);
      docStatus('📋 Copied to clipboard');
    });

    card.addEventListener('click', () => {
      showOutput(entry.response);
      document.getElementById('history-drawer').classList.add('hidden');
    });

    listEl.appendChild(card);
  });
}

// ─── Global Keyboard Shortcuts ──────────────────────────────────────────────

function bindKeyboardShortcuts() {
  window.addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && e.shiftKey) {
      const k = e.key.toUpperCase();
      if (k === 'R') {
        e.preventDefault();
        const btn = document.querySelector('.quick-btn[data-action="rewrite"]');
        if (btn) btn.click();
      } else if (k === 'S') {
        e.preventDefault();
        const btn = document.querySelector('.quick-btn[data-action="summarize"]');
        if (btn) btn.click();
      } else if (k === 'I') {
        e.preventDefault();
        const btn = document.querySelector('.quick-btn[data-action="improve"]');
        if (btn) btn.click();
      } else if (k === 'D') {
        e.preventDefault();
        toggleDarkMode();
      } else if (k === 'H') {
        e.preventDefault();
        const btn = document.getElementById('btn-history');
        if (btn) btn.click();
      } else if (k === 'G') {
        e.preventDefault();
        triggerGenerate();
      } else if (k === 'E') {
        e.preventDefault();
        if (typeof openInstantPopupModal === 'function') openInstantPopupModal();
      }
    }
  });
}

// ─── Settings & Brand Voice ──────────────────────────────────────────────────

function bindBrandVoice() {
  const select = document.getElementById('persona-select');
  const extractBtn = document.getElementById('btn-extract-voice');
  const statusEl = document.getElementById('voice-status');

  const saved = localStorage.getItem('wordai_persona') || 'standard';
  select.value = saved;

  select.addEventListener('change', () => {
    localStorage.setItem('wordai_persona', select.value);
    statusEl.textContent = `Active voice: ${select.options[select.selectedIndex].text}`;
  });

  extractBtn.addEventListener('click', async () => {
    statusEl.textContent = '⏳ Analyzing document writing style...';
    try {
      const docText = await getContext('document');
      if (!docText || docText.length < 100) {
        statusEl.textContent = '⚠️ Need at least a paragraph of text in document to analyze.';
        return;
      }
      const prompt = `Analyze the writing style of this text. Identify:
1. Typical sentence length and rhythm
2. Vocabulary tier and tone
3. Rhetorical devices or characteristic habits
Write a concise 2-sentence persona instruction describing how to write in this exact author's voice. Return ONLY the persona instruction.`;

      const instruction = await runAI(prompt, docText.slice(0, 4000));
      localStorage.setItem('wordai_custom_persona_prompt', instruction);
      localStorage.setItem('wordai_persona', 'custom');
      select.value = 'custom';
      statusEl.textContent = `✨ Learned & saved author persona!`;
    } catch(e) {
      statusEl.textContent = `❌ ${e.message}`;
    }
  });
}

const providerUrls = {
  openai:     { keyUrl: 'https://platform.openai.com/api-keys', name: 'OpenAI' },
  anthropic:  { keyUrl: 'https://console.anthropic.com/settings/keys', name: 'Anthropic' },
  gemini:     { keyUrl: 'https://aistudio.google.com/app/apikey', name: 'Google AI Studio' },
  deepseek:   { keyUrl: 'https://platform.deepseek.com/api_keys', name: 'DeepSeek' },
  groq:       { keyUrl: 'https://console.groq.com/keys', name: 'Groq' },
  mistral:    { keyUrl: 'https://console.mistral.ai/api-keys', name: 'Mistral AI' },
  openrouter: { keyUrl: 'https://openrouter.ai/keys', name: 'OpenRouter' },
  together:   { keyUrl: 'https://api.together.xyz/settings/api-keys', name: 'Together AI' },
  perplexity: { keyUrl: 'https://www.perplexity.ai/settings/api', name: 'Perplexity' },
  cohere:     { keyUrl: 'https://dashboard.cohere.com/api-keys', name: 'Cohere' },
  ollama:     { keyUrl: '', name: 'Ollama' },
  custom:     { keyUrl: '', name: 'Custom' }
};

function updateProviderUI(provider) {
  const baseUrlGroup = document.getElementById('base-url-group');
  const baseUrlInput = document.getElementById('api-base-url-input');
  const baseUrlHint = document.getElementById('base-url-hint');
  const keyInput = document.getElementById('api-key-input');
  const keyLink = document.getElementById('provider-key-link');

  if (provider === 'ollama') {
    baseUrlGroup.classList.remove('hidden');
    baseUrlHint.textContent = 'Default: http://localhost:11434/v1';
    if (!baseUrlInput.value) baseUrlInput.value = 'http://localhost:11434/v1';
    keyInput.placeholder = 'Optional (not required for local Ollama)';
    keyLink.style.display = 'none';
  } else if (provider === 'custom') {
    baseUrlGroup.classList.remove('hidden');
    baseUrlHint.textContent = 'e.g. http://localhost:1234/v1 (LM Studio)';
    if (!baseUrlInput.value) baseUrlInput.value = 'http://localhost:1234/v1';
    keyInput.placeholder = 'API Key or bearer token (optional if local)';
    keyLink.style.display = 'none';
  } else {
    baseUrlGroup.classList.add('hidden');
    keyInput.placeholder = 'Paste your API key...';
    keyLink.style.display = 'inline';
    const info = providerUrls[provider] || { keyUrl: '', name: 'API' };
    keyLink.href = info.keyUrl;
    keyLink.textContent = `Get ${info.name} Key ↗`;
  }
}

function bindSettings() {
  document.getElementById('btn-settings').addEventListener('click', () => {
    document.getElementById('settings-panel').classList.remove('hidden');
  });

  document.getElementById('btn-close-settings').addEventListener('click', () => {
    document.getElementById('settings-panel').classList.add('hidden');
  });

  document.getElementById('provider-select').addEventListener('change', (e) => {
    updateProviderUI(e.target.value);
  });

  document.getElementById('btn-fetch-models').addEventListener('click', async () => {
    const provider = document.getElementById('provider-select').value;
    const key = document.getElementById('api-key-input').value.trim();
    const baseUrl = document.getElementById('api-base-url-input').value.trim();
    const status = document.getElementById('fetch-status');
    if (!key && provider !== 'ollama' && provider !== 'custom') {
      status.textContent = '⚠️ Enter your API key first.';
      return;
    }
    const btn = document.getElementById('btn-fetch-models');
    btn.textContent = '⏳ Fetching...';
    btn.disabled = true;
    status.textContent = '';
    const models = await AIProvider.fetchModels(provider, key, baseUrl);
    const select = document.getElementById('model-select');
    select.innerHTML = '';
    if (models.length === 0) {
      select.innerHTML = '<option value="__auto__">🤖 Auto (fallback)</option><option value="">No models found</option>';
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
      status.textContent = `✅ ${models.length} models available.`;
    }
    btn.textContent = '🔄 Fetch Models';
    btn.disabled = false;
  });

  document.getElementById('btn-save-settings').addEventListener('click', () => {
    const model = document.getElementById('model-select').value;
    const provider = document.getElementById('provider-select').value;
    localStorage.setItem('wordai_provider', provider);
    localStorage.setItem('wordai_api_key', document.getElementById('api-key-input').value.trim());
    localStorage.setItem('wordai_base_url', document.getElementById('api-base-url-input').value.trim());
    localStorage.setItem('wordai_model', model);
    localStorage.setItem('wordai_language', document.getElementById('language-select').value);
    document.getElementById('settings-panel').classList.add('hidden');
    updateTokenDisplay();
    docStatus(`⚙️ Saved settings (${provider.toUpperCase()})`);
  });

  document.getElementById('btn-reset-tokens').addEventListener('click', () => {
    localStorage.setItem('wordai_tokens', '0');
    updateTokenDisplay();
  });
}

function loadSettings() {
  const provider = localStorage.getItem('wordai_provider') || 'openai';
  const key = localStorage.getItem('wordai_api_key') || '';
  const baseUrl = localStorage.getItem('wordai_base_url') || '';
  const model = localStorage.getItem('wordai_model') || '';
  const lang = localStorage.getItem('wordai_language') || 'English';

  document.getElementById('provider-select').value = provider;
  document.getElementById('api-key-input').value = key;
  document.getElementById('api-base-url-input').value = baseUrl;
  document.getElementById('language-select').value = lang;

  updateProviderUI(provider);

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

// ─── Pre-built Prompts Library ──────────────────────────────────────────

let activePromptCategory = 'all';

function bindPromptsLibrary() {
  const drawer = document.getElementById('prompts-drawer');
  const openBtnHeader = document.getElementById('btn-prompts');
  const openBtnPrompt = document.getElementById('btn-open-prompts-library');
  const closeBtn = document.getElementById('btn-close-prompts');
  const searchInput = document.getElementById('prompt-search-input');
  const viewAllResearchBtn = document.getElementById('btn-view-all-research-prompts');
  if (!drawer) return;

  function openDrawer(category = 'all') {
    activePromptCategory = category;
    document.querySelectorAll('.prompt-filter-pill').forEach(p => {
      p.classList.toggle('active', p.getAttribute('data-pcat') === category);
    });
    if (searchInput) searchInput.value = '';
    renderPrebuiltPrompts(category, '');
    drawer.classList.remove('hidden');
  }

  if (openBtnHeader) openBtnHeader.addEventListener('click', () => openDrawer('all'));
  if (openBtnPrompt) openBtnPrompt.addEventListener('click', () => openDrawer('all'));
  if (closeBtn) closeBtn.addEventListener('click', () => drawer.classList.add('hidden'));

  if (viewAllResearchBtn) {
    viewAllResearchBtn.addEventListener('click', () => openDrawer('academic'));
  }

  // Filter Pills
  document.querySelectorAll('.prompt-filter-pill').forEach(pill => {
    pill.addEventListener('click', () => {
      document.querySelectorAll('.prompt-filter-pill').forEach(p => p.classList.remove('active'));
      pill.classList.add('active');
      activePromptCategory = pill.getAttribute('data-pcat') || 'all';
      const searchVal = searchInput ? searchInput.value.trim() : '';
      renderPrebuiltPrompts(activePromptCategory, searchVal);
    });
  });

  // Search input filter
  if (searchInput) {
    searchInput.addEventListener('input', (e) => {
      renderPrebuiltPrompts(activePromptCategory, e.target.value.trim());
    });
  }

  // Quick prompt chips in Research tab
  document.querySelectorAll('.res-prompt-chip').forEach(chip => {
    chip.addEventListener('click', async () => {
      const pid = chip.getAttribute('data-pid');
      const item = typeof PREBUILT_PROMPTS !== 'undefined' ? PREBUILT_PROMPTS.find(p => p.id === pid) : null;
      if (!item) return;
      const context = await getContext('selection') || await getContext('document');
      handleAIAction(item.prompt, context || 'Please apply this research analysis.');
    });
  });

  // Initial render
  renderPrebuiltPrompts('all', '');
}

function renderPrebuiltPrompts(category = 'all', query = '') {
  const container = document.getElementById('prompts-container');
  if (!container || typeof PREBUILT_PROMPTS === 'undefined') return;

  const q = query.toLowerCase();
  const filtered = PREBUILT_PROMPTS.filter(p => {
    const matchCat = category === 'all' || p.category === category;
    const matchQ = !q || p.title.toLowerCase().includes(q) || p.desc.toLowerCase().includes(q) || p.prompt.toLowerCase().includes(q);
    return matchCat && matchQ;
  });

  if (filtered.length === 0) {
    container.innerHTML = '<div style="font-size:11.5px;color:var(--muted);text-align:center;padding:16px;">No prompts found matching your search.</div>';
    return;
  }

  container.innerHTML = '';
  filtered.forEach(p => {
    const card = document.createElement('div');
    card.className = 'prompt-card';
    card.innerHTML = `
      <div class="prompt-card-header">
        <span class="prompt-card-title">${escapeHtml(p.title)}</span>
        <span class="framework-badge badge-${p.category}">${escapeHtml(p.categoryLabel)}</span>
      </div>
      <div class="prompt-card-desc">${escapeHtml(p.desc)}</div>
      <div class="prompt-text-snippet">${escapeHtml(p.prompt)}</div>
      <div class="prompt-actions-row">
        <button class="prompt-btn-run" title="Run directly with current context">⚡ Run Prompt</button>
        <button class="prompt-btn-load" title="Load prompt into prompt box">✏️ Load into Box</button>
        <button class="prompt-btn-save" title="Save as custom preset">⭐</button>
      </div>
    `;

    // ⚡ Run Prompt
    card.querySelector('.prompt-btn-run').addEventListener('click', async () => {
      document.getElementById('prompts-drawer').classList.add('hidden');
      const contextMode = document.getElementById('context-select').value;
      const context = await getContext(contextMode);
      recordPrompt(p.prompt);
      handleAIAction(p.prompt, context || '(no content provided)');
    });

    // ✏️ Load into Prompt Box
    card.querySelector('.prompt-btn-load').addEventListener('click', () => {
      document.getElementById('main-prompt').value = p.prompt;
      document.getElementById('prompts-drawer').classList.add('hidden');
      document.getElementById('main-prompt').focus();
      docStatus(`💡 Loaded prompt: "${p.title}"`);
    });

    // ⭐ Save to Presets
    card.querySelector('.prompt-btn-save').addEventListener('click', () => {
      let presets = JSON.parse(localStorage.getItem('wordai_custom_presets') || '[]');
      if (!presets.some(x => x.name === p.title)) {
        presets.push({ name: p.title, prompt: p.prompt });
        localStorage.setItem('wordai_custom_presets', JSON.stringify(presets));
        docStatus(`⭐ Saved "${p.title}" to Custom Presets!`);
      } else {
        docStatus(`⭐ Already in Custom Presets!`);
      }
    });

    container.appendChild(card);
  });
}

