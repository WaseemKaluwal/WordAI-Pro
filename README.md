# WordAI Pro — Microsoft Word Add-in

AI writing assistant that lives inside Word. No tab switching.

## Features (MVP v1)
- Quick actions: Rewrite, Summarize, Improve, Translate, Expand, Shorten
- Generate: Emails, Reports, Proposals, Outlines, Tables, Bullets
- Edit: Grammar check, Fix All, Clarity, Passive voice, Writing score
- Tone changer: Professional, Academic, Friendly, Persuasive, and more
- Document Chat: Ask AI about your entire document
- Multi-provider: OpenAI, Gemini, OpenRouter, Groq (BYOK)
- Language support: English, Urdu, Arabic, Spanish, French, and more
- Token usage tracker

## Setup

### 1. Install dependencies
```
npm install
```

### 2. Start the dev server
```
npm start
```
This opens Word and sideloads the add-in automatically.

### Manual Sideloading (Windows)
1. Open Word → File → Options → Trust Center → Trust Center Settings
2. Trusted Add-in Catalogs → add `https://localhost:3000`
3. Or use: Insert → Get Add-ins → My Add-ins → Upload My Add-in → select `manifest.xml`

### 3. Add your API key
- Click ⚙️ Settings in the task pane
- Choose your AI provider (OpenAI, Gemini, OpenRouter, Groq)
- Paste your API key
- Optionally set a custom model name
- Save

## File Structure
```
WordAI Pro/
├── manifest.xml       # Office Add-in manifest
├── taskpane.html      # Main UI
├── taskpane.css       # Styles
├── taskpane.js        # Core logic + Word API
├── ai-provider.js     # Multi-provider AI layer
├── commands.html      # Required by manifest
└── package.json       # Dev dependencies
```

## Supported AI Providers
| Provider    | Default Model       | Get Key |
|-------------|---------------------|---------|
| OpenAI      | gpt-4o              | platform.openai.com |
| Gemini      | gemini-1.5-flash    | aistudio.google.com |
| OpenRouter  | openai/gpt-4o       | openrouter.ai |
| Groq        | llama3-70b-8192     | console.groq.com |

## Roadmap
- v2: Web research, PDF analysis, Citations, Student/Business modes
- v3: AI agents, Templates marketplace, Google Docs integration
