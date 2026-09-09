// ai-provider.js — Multi-provider AI layer supporting 12 leading AI providers & local models

const AIProvider = (() => {

  const PROVIDERS = {
    openai: {
      name: 'OpenAI',
      url: () => 'https://api.openai.com/v1/chat/completions',
      defaultModel: 'gpt-4o',
      buildHeaders: (key) => ({ 'Content-Type': 'application/json', 'Authorization': `Bearer ${key}` }),
      buildBody: (model, messages) => ({ model, messages, temperature: 0.7 }),
      extractText: (data) => data.choices?.[0]?.message?.content || '',
      extractTokens: (data) => data.usage?.total_tokens || 0,
    },
    anthropic: {
      name: 'Anthropic Claude',
      url: () => 'https://api.anthropic.com/v1/messages',
      defaultModel: 'claude-3-5-sonnet-20241022',
      buildHeaders: (key) => ({
        'Content-Type': 'application/json',
        'x-api-key': key,
        'anthropic-version': '2023-06-01',
        'anthropic-dangerous-direct-browser-access': 'true',
      }),
      buildBody: (model, messages) => {
        const sys = messages.filter(m => m.role === 'system').map(m => m.content).join('\n\n');
        const filtered = messages
          .filter(m => m.role !== 'system')
          .map(m => {
            if (Array.isArray(m.content)) {
              const contentParts = m.content.map(p => {
                if (p.type === 'image_url') {
                  const url = p.image_url.url;
                  const commaIdx = url.indexOf(',');
                  const mimeMatch = url.match(/data:([^;]+);base64/);
                  const media_type = mimeMatch ? mimeMatch[1] : 'image/jpeg';
                  const data = commaIdx > -1 ? url.slice(commaIdx + 1) : url;
                  return { type: 'image', source: { type: 'base64', media_type, data } };
                }
                return { type: 'text', text: p.text || '' };
              });
              return { role: m.role, content: contentParts };
            }
            return { role: m.role, content: m.content };
          });
        return {
          model,
          max_tokens: 4096,
          system: sys || undefined,
          messages: filtered,
          temperature: 0.7,
        };
      },
      extractText: (data) => data.content?.find(c => c.type === 'text')?.text || data.content?.[0]?.text || '',
      extractTokens: (data) => (data.usage?.input_tokens || 0) + (data.usage?.output_tokens || 0),
    },
    gemini: {
      name: 'Google Gemini',
      url: (model, key) => `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${key}`,
      defaultModel: 'gemini-2.0-flash',
      buildHeaders: () => ({ 'Content-Type': 'application/json' }),
      buildBody: (model, messages) => ({
        contents: messages
          .filter(m => m.role !== 'system')
          .map(m => {
            if (Array.isArray(m.content)) {
              const parts = [];
              m.content.forEach(c => {
                if (c.type === 'text') parts.push({ text: c.text });
                if (c.type === 'image_url') {
                  const b64 = c.image_url.url.split(',')[1] || c.image_url.url;
                  const mime = c.image_url.url.includes('image/png') ? 'image/png' : 'image/jpeg';
                  parts.push({ inlineData: { mimeType: mime, data: b64 } });
                }
              });
              return { role: m.role === 'assistant' ? 'model' : 'user', parts };
            }
            return {
              role: m.role === 'assistant' ? 'model' : 'user',
              parts: [{ text: typeof m.content === 'string' ? m.content : JSON.stringify(m.content) }],
            };
          }),
        systemInstruction: messages.find(m => m.role === 'system')
          ? { parts: [{ text: messages.find(m => m.role === 'system').content }] }
          : undefined,
      }),
      extractText: (data) => data.candidates?.[0]?.content?.parts?.[0]?.text || '',
      extractTokens: (data) => data.usageMetadata?.totalTokenCount || 0,
    },
    deepseek: {
      name: 'DeepSeek',
      url: () => 'https://api.deepseek.com/chat/completions',
      defaultModel: 'deepseek-chat',
      buildHeaders: (key) => ({ 'Content-Type': 'application/json', 'Authorization': `Bearer ${key}` }),
      buildBody: (model, messages) => ({
        model,
        messages: messages.map(m => Array.isArray(m.content)
          ? { role: m.role, content: m.content.filter(p => p.type === 'text').map(p => p.text).join('\n') }
          : m),
        temperature: 0.7,
      }),
      extractText: (data) => data.choices?.[0]?.message?.content || '',
      extractTokens: (data) => data.usage?.total_tokens || 0,
    },
    groq: {
      name: 'Groq',
      url: () => 'https://api.groq.com/openai/v1/chat/completions',
      defaultModel: 'llama-3.3-70b-versatile',
      buildHeaders: (key) => ({ 'Content-Type': 'application/json', 'Authorization': `Bearer ${key}` }),
      buildBody: (model, messages) => ({ model, messages, temperature: 0.7 }),
      extractText: (data) => data.choices?.[0]?.message?.content || '',
      extractTokens: (data) => data.usage?.total_tokens || 0,
    },
    mistral: {
      name: 'Mistral AI',
      url: () => 'https://api.mistral.ai/v1/chat/completions',
      defaultModel: 'mistral-large-latest',
      buildHeaders: (key) => ({ 'Content-Type': 'application/json', 'Authorization': `Bearer ${key}` }),
      buildBody: (model, messages) => ({
        model,
        messages: messages.map(m => Array.isArray(m.content)
          ? { role: m.role, content: m.content.filter(p => p.type === 'text').map(p => p.text).join('\n') }
          : m),
        temperature: 0.7,
      }),
      extractText: (data) => data.choices?.[0]?.message?.content || '',
      extractTokens: (data) => data.usage?.total_tokens || 0,
    },
    openrouter: {
      name: 'OpenRouter',
      url: () => 'https://openrouter.ai/api/v1/chat/completions',
      defaultModel: 'openai/gpt-4o',
      buildHeaders: (key) => ({
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${key}`,
        'HTTP-Referer': 'https://github.com/waseemkaluwal/wordai-pro',
        'X-Title': 'WordAI Pro',
      }),
      buildBody: (model, messages) => ({ model, messages }),
      extractText: (data) => data.choices?.[0]?.message?.content || '',
      extractTokens: (data) => data.usage?.total_tokens || 0,
    },
    together: {
      name: 'Together AI',
      url: () => 'https://api.together.xyz/v1/chat/completions',
      defaultModel: 'meta-llama/Llama-3.3-70B-Instruct-Turbo',
      buildHeaders: (key) => ({ 'Content-Type': 'application/json', 'Authorization': `Bearer ${key}` }),
      buildBody: (model, messages) => ({ model, messages, temperature: 0.7 }),
      extractText: (data) => data.choices?.[0]?.message?.content || '',
      extractTokens: (data) => data.usage?.total_tokens || 0,
    },
    perplexity: {
      name: 'Perplexity AI',
      url: () => 'https://api.perplexity.ai/chat/completions',
      defaultModel: 'sonar',
      buildHeaders: (key) => ({ 'Content-Type': 'application/json', 'Authorization': `Bearer ${key}` }),
      buildBody: (model, messages) => ({
        model,
        messages: messages.map(m => Array.isArray(m.content)
          ? { role: m.role, content: m.content.filter(p => p.type === 'text').map(p => p.text).join('\n') }
          : m),
        temperature: 0.7,
      }),
      extractText: (data) => data.choices?.[0]?.message?.content || '',
      extractTokens: (data) => data.usage?.total_tokens || 0,
    },
    cohere: {
      name: 'Cohere',
      url: () => 'https://api.cohere.com/v2/chat',
      defaultModel: 'command-r-plus-08-2024',
      buildHeaders: (key) => ({ 'Content-Type': 'application/json', 'Authorization': `Bearer ${key}` }),
      buildBody: (model, messages) => ({
        model,
        messages: messages.map(m => ({
          role: m.role,
          content: Array.isArray(m.content) ? m.content.filter(p => p.type === 'text').map(p => p.text).join('\n') : m.content,
        })),
      }),
      extractText: (data) => data.message?.content?.[0]?.text || data.text || '',
      extractTokens: (data) => (data.usage?.tokens?.input_tokens || 0) + (data.usage?.tokens?.output_tokens || 0),
    },
    ollama: {
      name: 'Ollama (Local / Offline)',
      url: (_m, _k, baseUrl) => `${(baseUrl || 'http://localhost:11434/v1').replace(/\/+$/, '')}/chat/completions`,
      defaultModel: 'llama3.2',
      buildHeaders: (key) => ({ 'Content-Type': 'application/json', 'Authorization': `Bearer ${key || 'ollama'}` }),
      buildBody: (model, messages) => ({ model, messages, temperature: 0.7 }),
      extractText: (data) => data.choices?.[0]?.message?.content || '',
      extractTokens: (data) => data.usage?.total_tokens || 0,
    },
    custom: {
      name: 'Custom (OpenAI-Compatible)',
      url: (_m, _k, baseUrl) => `${(baseUrl || 'http://localhost:1234/v1').replace(/\/+$/, '')}/chat/completions`,
      defaultModel: 'default',
      buildHeaders: (key) => ({ 'Content-Type': 'application/json', 'Authorization': `Bearer ${key || 'custom'}` }),
      buildBody: (model, messages) => ({ model, messages, temperature: 0.7 }),
      extractText: (data) => data.choices?.[0]?.message?.content || '',
      extractTokens: (data) => data.usage?.total_tokens || 0,
    },
  };

  const FALLBACKS = {
    openai:     ['gpt-4o', 'gpt-4o-mini', 'o3-mini', 'o1-mini', 'gpt-4-turbo', 'gpt-3.5-turbo'],
    anthropic:  ['claude-3-5-sonnet-20241022', 'claude-3-5-haiku-20241022', 'claude-3-opus-20240229', 'claude-3-sonnet-20240229', 'claude-3-haiku-20240307'],
    gemini:     ['gemini-2.0-flash', 'gemini-1.5-flash', 'gemini-1.5-pro', 'gemini-pro'],
    deepseek:   ['deepseek-chat', 'deepseek-reasoner'],
    groq:       ['llama-3.3-70b-versatile', 'llama-3.1-8b-instant', 'llama3-70b-8192', 'llama3-8b-8192', 'mixtral-8x7b-32768'],
    mistral:    ['mistral-large-latest', 'mistral-small-latest', 'codestral-latest', 'open-mistral-nemo', 'open-mixtral-8x22b'],
    openrouter: ['openai/gpt-4o', 'anthropic/claude-3.5-sonnet', 'deepseek/deepseek-r1', 'google/gemini-2.0-flash-001', 'meta-llama/llama-3.3-70b-instruct'],
    together:   ['meta-llama/Llama-3.3-70B-Instruct-Turbo', 'deepseek-ai/DeepSeek-R1', 'deepseek-ai/DeepSeek-V3', 'meta-llama/Meta-Llama-3.1-70B-Instruct-Turbo', 'Qwen/Qwen2.5-72B-Instruct-Turbo'],
    perplexity: ['sonar', 'sonar-pro', 'sonar-reasoning'],
    cohere:     ['command-r-plus-08-2024', 'command-r-08-2024', 'command-r-plus', 'command-r'],
    ollama:     ['llama3.2', 'llama3.1', 'deepseek-r1', 'mistral', 'qwen2.5', 'phi3'],
    custom:     ['default'],
  };

  async function callModel(p, model, apiKey, messages, baseUrl) {
    const url = typeof p.url === 'function' ? p.url(model, apiKey, baseUrl) : p.url;
    const res = await fetch(url, {
      method: 'POST',
      headers: p.buildHeaders(apiKey),
      body: JSON.stringify(p.buildBody(model, messages)),
    });
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      const msg = err.error?.message || err.message || err.detail || `API error ${res.status}`;
      throw new Error(msg);
    }
    const data = await res.json();
    return { text: p.extractText(data), tokens: p.extractTokens(data) };
  }

  async function call(messages) {
    const providerName = localStorage.getItem('wordai_provider') || 'openai';
    const apiKey = localStorage.getItem('wordai_api_key') || (providerName === 'ollama' ? 'ollama' : '');
    const baseUrl = localStorage.getItem('wordai_base_url') || (providerName === 'ollama' ? 'http://localhost:11434/v1' : 'http://localhost:1234/v1');
    const savedModel = localStorage.getItem('wordai_model') || '';

    if (!apiKey && providerName !== 'ollama') {
      throw new Error(`No API key set for ${providerName.toUpperCase()}. Open Settings ⚙️ to add your key.`);
    }

    const p = PROVIDERS[providerName];
    if (!p) throw new Error(`Unknown provider: ${providerName}`);

    // Build model list to try: saved/default first, then fallbacks
    const isAuto = savedModel === '__auto__' || !savedModel;
    const fallbacks = FALLBACKS[providerName] || [];
    const modelsToTry = isAuto
      ? (fallbacks.length ? fallbacks : [p.defaultModel])
      : [savedModel, ...fallbacks.filter(m => m !== savedModel)];

    let lastError;
    for (const model of modelsToTry) {
      try {
        const { text, tokens } = await callModel(p, model, apiKey, messages, baseUrl);
        if (isAuto) localStorage.setItem('wordai_working_model', model);
        const prev = parseInt(localStorage.getItem('wordai_tokens') || '0', 10);
        localStorage.setItem('wordai_tokens', prev + tokens);
        return { text, tokens, model };
      } catch (err) {
        lastError = err;
        // Only fallback on model-not-found or quota errors, not auth errors
        if (err.message.includes('401') || err.message.includes('API key') || err.message.includes('Unauthorized') || err.message.includes('authentication')) {
          throw err;
        }
      }
    }
    throw new Error(`All models failed for ${p.name}. Last error: ${lastError?.message}`);
  }

  async function fetchModels(providerName, apiKey, baseUrl) {
    const fallbacks = FALLBACKS[providerName] || [];
    try {
      if (providerName === 'gemini') {
        const res = await fetch(`https://generativelanguage.googleapis.com/v1beta/models?key=${apiKey}`);
        if (res.ok) {
          const data = await res.json();
          const list = (data.models || [])
            .filter(m => m.supportedGenerationMethods?.includes('generateContent'))
            .map(m => m.name.replace('models/', ''));
          if (list.length) return list;
        }
      }

      if (providerName === 'openai') {
        const res = await fetch('https://api.openai.com/v1/models', {
          headers: { 'Authorization': `Bearer ${apiKey}` },
        });
        if (res.ok) {
          const data = await res.json();
          const list = (data.data || [])
            .filter(m => m.id.startsWith('gpt') || m.id.startsWith('o1') || m.id.startsWith('o3'))
            .map(m => m.id)
            .sort();
          if (list.length) return list;
        }
      }

      if (providerName === 'deepseek') {
        const res = await fetch('https://api.deepseek.com/models', {
          headers: { 'Authorization': `Bearer ${apiKey}` },
        });
        if (res.ok) {
          const data = await res.json();
          const list = (data.data || []).map(m => m.id).sort();
          if (list.length) return list;
        }
      }

      if (providerName === 'groq') {
        const res = await fetch('https://api.groq.com/openai/v1/models', {
          headers: { 'Authorization': `Bearer ${apiKey}` },
        });
        if (res.ok) {
          const data = await res.json();
          const list = (data.data || []).map(m => m.id).sort();
          if (list.length) return list;
        }
      }

      if (providerName === 'mistral') {
        const res = await fetch('https://api.mistral.ai/v1/models', {
          headers: { 'Authorization': `Bearer ${apiKey}` },
        });
        if (res.ok) {
          const data = await res.json();
          const list = (data.data || []).map(m => m.id).sort();
          if (list.length) return list;
        }
      }

      if (providerName === 'openrouter') {
        const res = await fetch('https://openrouter.ai/api/v1/models', {
          headers: { 'Authorization': `Bearer ${apiKey}` },
        });
        if (res.ok) {
          const data = await res.json();
          const list = (data.data || []).map(m => m.id).sort();
          if (list.length) return list;
        }
      }

      if (providerName === 'together') {
        const res = await fetch('https://api.together.xyz/v1/models', {
          headers: { 'Authorization': `Bearer ${apiKey}` },
        });
        if (res.ok) {
          const data = await res.json();
          const list = (data || []).map(m => m.id || m.name).filter(Boolean).sort();
          if (list.length) return list;
        }
      }

      if (providerName === 'ollama') {
        const base = (baseUrl || 'http://localhost:11434/v1').replace(/\/+$/, '');
        try {
          const res = await fetch(`${base}/models`);
          if (res.ok) {
            const data = await res.json();
            const list = (data.data || []).map(m => m.id).filter(Boolean);
            if (list.length) return list;
          }
        } catch (_) {}
        try {
          const tagsUrl = `${base.replace(/\/v1\/?$/, '')}/api/tags`;
          const res = await fetch(tagsUrl);
          if (res.ok) {
            const data = await res.json();
            const list = (data.models || []).map(m => m.name).filter(Boolean);
            if (list.length) return list;
          }
        } catch (_) {}
      }

      if (providerName === 'custom') {
        const base = (baseUrl || 'http://localhost:1234/v1').replace(/\/+$/, '');
        const res = await fetch(`${base}/models`, {
          headers: apiKey ? { 'Authorization': `Bearer ${apiKey}` } : {},
        });
        if (res.ok) {
          const data = await res.json();
          const list = (data.data || []).map(m => m.id).filter(Boolean);
          if (list.length) return list;
        }
      }

    } catch (_) {
      // Return fallbacks on network error
    }
    return fallbacks;
  }

  return { call, fetchModels, PROVIDERS, FALLBACKS };
})();
