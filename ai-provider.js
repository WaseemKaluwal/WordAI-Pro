// ai-provider.js — Multi-provider AI layer

const AIProvider = (() => {

  const PROVIDERS = {
    openai: {
      url: 'https://api.openai.com/v1/chat/completions',
      defaultModel: 'gpt-4o',
      buildBody: (model, messages) => ({ model, messages, temperature: 0.7 }),
      buildHeaders: (key) => ({ 'Content-Type': 'application/json', 'Authorization': `Bearer ${key}` }),
      extractText: (data) => data.choices[0].message.content,
      extractTokens: (data) => data.usage?.total_tokens || 0,
    },
    gemini: {
      url: (model, key) => `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${key}`,
      defaultModel: 'gemini-2.0-flash',
      buildBody: (model, messages) => ({
        contents: messages
          .filter(m => m.role !== 'system')
          .map(m => ({
            role: m.role === 'assistant' ? 'model' : 'user',
            parts: [{ text: m.content }]
          })),
        systemInstruction: messages.find(m => m.role === 'system')
          ? { parts: [{ text: messages.find(m => m.role === 'system').content }] }
          : undefined,
      }),
      buildHeaders: () => ({ 'Content-Type': 'application/json' }),
      extractText: (data) => data.candidates[0].content.parts[0].text,
      extractTokens: (data) => (data.usageMetadata?.totalTokenCount) || 0,
    },
    openrouter: {
      url: 'https://openrouter.ai/api/v1/chat/completions',
      defaultModel: 'openai/gpt-4o',
      buildBody: (model, messages) => ({ model, messages }),
      buildHeaders: (key) => ({ 'Content-Type': 'application/json', 'Authorization': `Bearer ${key}` }),
      extractText: (data) => data.choices[0].message.content,
      extractTokens: (data) => data.usage?.total_tokens || 0,
    },
    groq: {
      url: 'https://api.groq.com/openai/v1/chat/completions',
      defaultModel: 'llama3-70b-8192',
      buildBody: (model, messages) => ({ model, messages, temperature: 0.7 }),
      buildHeaders: (key) => ({ 'Content-Type': 'application/json', 'Authorization': `Bearer ${key}` }),
      extractText: (data) => data.choices[0].message.content,
      extractTokens: (data) => data.usage?.total_tokens || 0,
    },
  };

  const FALLBACKS = {
    openai:     ['gpt-4o', 'gpt-4o-mini', 'gpt-4-turbo', 'gpt-4', 'gpt-3.5-turbo'],
    gemini:     ['gemini-2.0-flash', 'gemini-1.5-flash', 'gemini-1.5-pro', 'gemini-pro'],
    groq:       ['llama3-70b-8192', 'llama3-8b-8192', 'mixtral-8x7b-32768', 'gemma2-9b-it'],
    openrouter: ['openai/gpt-4o', 'openai/gpt-4o-mini', 'google/gemini-flash-1.5', 'meta-llama/llama-3-70b-instruct'],
  };

  async function callModel(p, model, apiKey, messages) {
    const url = typeof p.url === 'function' ? p.url(model, apiKey) : p.url;
    const res = await fetch(url, {
      method: 'POST',
      headers: p.buildHeaders(apiKey),
      body: JSON.stringify(p.buildBody(model, messages)),
    });
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err.error?.message || `API error ${res.status}`);
    }
    const data = await res.json();
    return { text: p.extractText(data), tokens: p.extractTokens(data) };
  }

  async function call(messages) {
    const providerName = localStorage.getItem('wordai_provider') || 'openai';
    const apiKey = localStorage.getItem('wordai_api_key') || '';
    const savedModel = localStorage.getItem('wordai_model') || '';

    if (!apiKey) throw new Error('No API key set. Open Settings ⚙️ to add your key.');

    const p = PROVIDERS[providerName];
    if (!p) throw new Error(`Unknown provider: ${providerName}`);

    // Build model list to try: saved/default first, then fallbacks
    const isAuto = savedModel === '__auto__' || !savedModel;
    const fallbacks = FALLBACKS[providerName] || [];
    const modelsToTry = isAuto
      ? fallbacks
      : [savedModel, ...fallbacks.filter(m => m !== savedModel)];

    let lastError;
    for (const model of modelsToTry) {
      try {
        const { text, tokens } = await callModel(p, model, apiKey, messages);
        // If auto mode and model changed, persist the working model
        if (isAuto) localStorage.setItem('wordai_working_model', model);
        const prev = parseInt(localStorage.getItem('wordai_tokens') || '0', 10);
        localStorage.setItem('wordai_tokens', prev + tokens);
        return { text, tokens, model };
      } catch (err) {
        lastError = err;
        // Only fallback on model-not-found or quota errors, not auth errors
        if (err.message.includes('401') || err.message.includes('API key')) throw err;
      }
    }
    throw new Error(`All models failed. Last error: ${lastError?.message}`);
  }

  async function fetchModels(providerName, apiKey) {
    try {
      if (providerName === 'gemini') {
        const res = await fetch(`https://generativelanguage.googleapis.com/v1beta/models?key=${apiKey}`);
        const data = await res.json();
        return (data.models || [])
          .filter(m => m.supportedGenerationMethods?.includes('generateContent'))
          .map(m => m.name.replace('models/', ''));
      }
      if (providerName === 'openai') {
        const res = await fetch('https://api.openai.com/v1/models', {
          headers: { 'Authorization': `Bearer ${apiKey}` }
        });
        const data = await res.json();
        return (data.data || [])
          .filter(m => m.id.startsWith('gpt'))
          .map(m => m.id)
          .sort();
      }
      if (providerName === 'groq') {
        const res = await fetch('https://api.groq.com/openai/v1/models', {
          headers: { 'Authorization': `Bearer ${apiKey}` }
        });
        const data = await res.json();
        return (data.data || []).map(m => m.id).sort();
      }
      if (providerName === 'openrouter') {
        const res = await fetch('https://openrouter.ai/api/v1/models', {
          headers: { 'Authorization': `Bearer ${apiKey}` }
        });
        const data = await res.json();
        return (data.data || []).map(m => m.id).sort();
      }
    } catch (e) { return []; }
    return [];
  }

  return { call, fetchModels };
})();
