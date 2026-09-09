# ⚡ WordAI Pro — Microsoft Word AI Add-in

AI writing assistant + academic research copilot that lives inside Word. No tab switching.

🔗 **Live Add-in:** https://waseemkaluwal.github.io/wordai-pro/taskpane.html

---

## Features

### ✏️ Quick Actions
Rewrite, Summarize, Improve, Translate, Expand, Shorten — one click on selected text.

### ✨ Generate
Email, Report, Proposal, Cover Letter, Outline, Introduction, Conclusion, Bullets, Paragraphs, Table.

### 📝 Edit
Grammar Check, Fix All, Improve Clarity, Fix Passive Voice, Simplify Language, Writing Score.

### 🎭 Tone Changer
Professional, Academic, Friendly, Persuasive, Simple, Formal, Casual, Confident.

### 🤖 Document Chat
Ask AI anything about your entire document. Maintains conversation history.

### 🎓 Research Copilot (Academic Mode)
| Tool | What it does |
|---|---|
| 📋 Research Planner | Research questions, hypothesis, methodology, chapter structure |
| 📚 Literature Review | Themes, agreements, contradictions, research gaps |
| 🔍 Evidence Matrix | Maps claims → evidence type → where to find sources |
| ⚖️ Argument Checker | Flags unsupported claims, logic gaps, contradictions |
| 🎤 Oral Defense Prep | 10 examiner questions + what strong answers need |
| 📄 Write Abstract | Structured 250-300 word academic abstract |
| 🔬 Methodology | Full methodology section (design, sampling, validity, ethics) |
| 📖 Bibliography | Formats citations in APA / MLA / Chicago / Harvard / IEEE |

### 🔑 Multi-Provider AI (BYOK)
| Provider | Default Model | Get Key |
|---|---|---|
| OpenAI | gpt-4o | platform.openai.com |
| Google Gemini | gemini-2.0-flash | aistudio.google.com |
| OpenRouter | openai/gpt-4o | openrouter.ai |
| Groq | llama3-70b-8192 | console.groq.com |

- **Auto-fetch models** — click 🔄 Fetch in Settings to load all available models from your provider live
- **Token usage tracker** built in

### 🌍 Language Support
English, Urdu, Roman Urdu, Arabic, Spanish, French, German, Chinese, Hindi

---

## Installation (Word Online — Free)

### 1. Go to Word Online
- Open **office.com** → sign in with a free Microsoft account → open any document

### 2. Upload the manifest
- Insert → Add-ins → My Add-ins → **Upload My Add-in**
- Upload `manifest.xml`

### 3. Add your API key
- Click ⚙️ Settings in the task pane
- Choose your AI provider
- Paste your API key
- Click **🔄 Fetch** to load available models
- Select a model → Save

---

## Installation (Word Desktop — requires licensed Word)

```
npm install
npx office-addin-dev-certs install
npm start
```

This starts a local server and sideloads the add-in into Word automatically.

---

## File Structure

```
WordAI Pro/
├── manifest.xml       # Office Add-in manifest (points to GitHub Pages)
├── taskpane.html      # Main UI
├── taskpane.css       # Styles
├── taskpane.js        # Core logic + Word API
├── ai-provider.js     # Multi-provider AI layer + model fetching
├── commands.html      # Required by manifest
└── package.json       # Dev dependencies
```

---

## Roadmap
- v2: Web research, PDF analysis, Citations, Student/Business modes
- v3: AI agents, Templates marketplace, Google Docs integration
