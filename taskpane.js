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
}

if (typeof Office !== 'undefined') {
  Office.onReady(() => initApp());
} else {
  document.addEventListener('DOMContentLoaded', initApp);
}

// ─── Helpers ────────────────────────────────────────────────────────────────

function showLoading(show) {
  document.getElementById('loading').classList.toggle('hidden', !show);
}

function showOutput(text) {
  const section = document.getElementById('output-section');
  const el = document.getElementById('output-text');
  el.textContent = text;
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

async function insertText(text, replace = false) {
  if (typeof Word === 'undefined') {
    showError('Word API not available in this context.');
    return;
  }
  return Word.run(async (ctx) => {
    const sel = ctx.document.getSelection();
    if (replace) {
      sel.insertText(text, Word.InsertLocation.replace);
    } else {
      sel.insertText('\n' + text, Word.InsertLocation.after);
    }
    await ctx.sync();
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
      const docLine = docText ? `\n\nDocument content:\n${docText.slice(0, 6000)}` : '';

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

      const userContent = [topicLine, docLine].filter(Boolean).join('') || 'No topic or document provided. Please enter a research topic above.';
      await handleAIAction(prompts[action], userContent);
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

// ─── Output Actions ──────────────────────────────────────────────────────────

function bindOutputActions() {
  document.getElementById('btn-insert').addEventListener('click', async () => {
    const text = document.getElementById('output-text').textContent;
    await insertText(text, false);
  });

  document.getElementById('btn-replace').addEventListener('click', async () => {
    const text = document.getElementById('output-text').textContent;
    await insertText(text, true);
  });

  document.getElementById('btn-copy').addEventListener('click', () => {
    const text = document.getElementById('output-text').textContent;
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

  document.getElementById('btn-save-settings').addEventListener('click', () => {
    localStorage.setItem('wordai_provider', document.getElementById('provider-select').value);
    localStorage.setItem('wordai_api_key', document.getElementById('api-key-input').value);
    localStorage.setItem('wordai_model', document.getElementById('model-input').value);
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
  document.getElementById('model-input').value = model;
  document.getElementById('language-select').value = lang;
  updateTokenDisplay();
}
