# ⚡ WordAI Pro — Microsoft Word AI Copilot & Research Assistant

[![Office Add-in](https://img.shields.io/badge/Office_Add--in-Word_2016%2B_|_Online_|_Mac_|_iPad-0078d4?style=for-the-badge&logo=microsoftword&logoColor=white)](https://waseemkaluwal.github.io/wordai-pro/)
[![Manifest Validated](https://img.shields.io/badge/Manifest-100%25_Validated-10b981?style=for-the-badge&logo=checkmarx&logoColor=white)](manifest.xml)
[![AI Providers](https://img.shields.io/badge/AI_Providers-12_Supported_(BYOK)-6366f1?style=for-the-badge&logo=openai&logoColor=white)](#-multi-provider-ai-suite-12-providers--byok)
[![Local LLM Support](https://img.shields.io/badge/Local_AI-Ollama_|_LM_Studio-f97316?style=for-the-badge&logo=ollama&logoColor=white)](#-privacy--100-local-offline-ai-ollama)
[![Privacy First](https://img.shields.io/badge/Privacy-100%25_Client--Side-14b8a6?style=for-the-badge&logo=shield&logoColor=white)](#-privacy-architecture--security)
[![License](https://img.shields.io/badge/License-Non--Commercial_(No--Resale)-8b5cf6?style=for-the-badge)](LICENSE)

**WordAI Pro** is an open-source, full-featured **AI Writing Assistant, Academic Research Copilot, and Document Automation Studio** that operates directly inside **Microsoft Word** (Word Online, Word for Windows Desktop, Word for Mac, and iPad). 

Equipped with direct page canvas in-place editing, a native rich Word styling engine, multi-model support across **12 leading AI providers** (including GPT-4o, Claude 3.5 Sonnet, Gemini 2.0 Flash, DeepSeek R1, and local offline Ollama), **28 pre-built document frameworks**, **48 curated prompts**, live citation discovery, and visual redline diff inspection—all with zero external servers and zero telemetry (100% client-side BYOK).

<p align="center">
  <img src="https://waseemkaluwal.github.io/wordai-pro/screenshots/screenshot-hero.png" alt="WordAI Pro running live inside Microsoft Word Desktop" width="100%" />
</p>

---

### 🌐 Live Production Deployment
- **Live Taskpane URL:** [https://waseemkaluwal.github.io/wordai-pro/taskpane.html](https://waseemkaluwal.github.io/wordai-pro/taskpane.html)
- **Manifest File:** [`manifest.xml`](manifest.xml) (Ready for 1-click sideloading into Microsoft Word)
- **Interactive In-Place Copilot:** Press `Ctrl + Shift + E` inside Microsoft Word

---

## 📑 Table of Contents
1. [Why WordAI Pro? (Comparison Matrix)](#-why-wordai-pro-vs-microsoft-copilot)
2. [Visual Tour & Interface Showcase](#-visual-tour--interface-showcase)
3. [Core Architectural Highlights](#-core-architectural-highlights)
   - [Direct In-Place Page Canvas Editing & Popup Copilot](#1-direct-in-place-page-canvas-editing--popup-copilot)
   - [Native Rich Word Formatting Engine (Zero Markdown Asterisks)](#2-native-rich-word-formatting-engine-zero-markdown-asterisks)
   - [Multi-Provider AI Suite (12 Providers — BYOK)](#3-multi-provider-ai-suite-12-providers--byok)
   - [Document Frameworks Studio (28 Pre-built Scaffolds)](#4-document-frameworks-studio-28-pre-built-scaffolds)
   - [Pre-built Strategic Prompts Studio (48 Expert Prompts)](#5-pre-built-strategic-prompts-studio-48-expert-prompts)
   - [Visual Redline Diff & Side-by-Side Evaluation](#6-visual-redline-diff--side-by-side-evaluation)
   - [The 6 Killer Features Suite (New in v1.0.1)](#7-the-6-killer-features-suite-new-in-v101)
3. [Complete 36-Feature Capabilities Matrix](#-complete-36-feature-capabilities-matrix)
4. [Keyboard Shortcuts Cheat Sheet](#-global-keyboard-shortcuts)
5. [Installation & Sideloading Guide](#-installation--setup-guide)
   - [Method 1: Word Online (Free — 60 Seconds)](#method-1-word-online-free--60-seconds)
   - [Method 2: Word Desktop (Windows / macOS)](#method-2-word-desktop-windows--macos)
   - [Deploying to GitHub Pages](#deploying-to-github-pages-single-root-folder)
6. [Privacy Architecture & Security](#-privacy-architecture--security)
7. [Frequently Asked Questions (FAQ)](#-frequently-asked-questions-faq)
8. [File Structure](#-repository-file-structure)
9. [Contributing & License](#-contributing--license)

---

## ⚡ Why WordAI Pro? vs. Microsoft Copilot

| Feature / Metric | Microsoft Copilot for M365 | WordAI Pro |
| :--- | :--- | :--- |
| **Subscription Cost** | **$30 / user / month** + M365 license | **100% Free (Non-Commercial)** |
| **AI Model Choice** | Locked to single OpenAI model | **12 Providers**: GPT-4o, Claude 3.5, Gemini 2.0, DeepSeek R1, Groq, Mistral, Ollama |
| **Local / Offline AI** | ❌ Not supported (Cloud only) | ✅ **100% Offline with Ollama & LM Studio** |
| **API Key Architecture** | Proprietary Microsoft billing | **BYOK (Bring Your Own Key)** — Pay fractions of a cent per prompt |
| **Chat With Document** | Basic cloud search | ✅ **Full Document Q&A with Live Word Sync & 1-Click Canvas Insertion** |
| **AI Humanizer Engine** | ❌ None (sounds robotic) | ✅ **Anti-AI Detector & Burstiness Polish with 98% Human Scorecard** |
| **Academic Citations** | Hallucinated text references | ✅ **250M+ Real Papers (OpenAlex & Crossref) + Auto-Bibliography at Doc End** |
| **Writing DNA Cloner** | ❌ Generic voice | ✅ **Personal Writing DNA Studio (Formality & Rhythm Profiling)** |
| **Direct In-Place Editing** | Limited canvas suggestions | ✅ **Live highlighted text rewrite (`Ctrl+Shift+E`)** |
| **Rich Word Formatting** | Inconsistent Markdown output | ✅ **Native Word HTML Engine (`<strong>`, Headings, Tables, Lists)** |
| **Text to Word Tables** | Plain text tables | ✅ **Native Styled HTML & Word Table Generation** |
| **Document Frameworks** | Generic empty drafts | ✅ **28 Full Industry Frameworks + Compliance Audit** |
| **Format-Safe Translator** | Strips formatting | ✅ **30+ Languages with Strict HTML & Markdown Tag Preservation** |
| **Visual Redline Diff** | Standard Word track changes | ✅ **Interactive side-by-side & LCS visual redline diff viewer** |
| **Data Privacy** | Enterprise cloud processing | ✅ **Zero server: All API calls go directly from your browser to the LLM** |

---

## 📸 Visual Tour & Interface Showcase

| 🎨 **AI Content Generation & In-Place Copilot** | 🔬 **Research Copilot & Academic Paper Discovery** |
| :---: | :---: |
| <img src="https://waseemkaluwal.github.io/wordai-pro/screenshots/screenshot-generate.png" width="100%" alt="AI Content Generation & In-Place Copilot" /> | <img src="https://waseemkaluwal.github.io/wordai-pro/screenshots/screenshot-research.png" width="100%" alt="Research Copilot & Academic Paper Discovery" /> |
| 📋 **28 Document Frameworks Studio** | ⚡ **48 Strategic Prompts Library** |
| <img src="https://waseemkaluwal.github.io/wordai-pro/screenshots/screenshot-frameworks.png" width="100%" alt="28 Document Frameworks Studio" /> | <img src="https://waseemkaluwal.github.io/wordai-pro/screenshots/screenshot-prompts.png" width="100%" alt="48 Strategic Prompts Library" /> |

---

## 💎 Core Architectural Highlights

### 1. Direct In-Place Page Canvas Editing & Popup Copilot
WordAI Pro writes, rewrites, and inserts **directly onto the Microsoft Word document canvas (page body)** in real time:
- **Instant Highlight Detection:** Detects highlighted text across Word in **380ms** via WordApi `context.document.onSelectionChanged`.
- **Frameless In-Place Popup Modal (`Ctrl + Shift + E`):** An elevated, frosted-glass popup dialog that animates over your workspace with the active text snippet, live word count, and direct instruction box.
- **7 One-Tap Transformation Chips:** `✏️ Rewrite`, `✨ Improve`, `🔍 Fix Grammar`, `➖ Shorten`, `➕ Expand`, `💼 Formal`, `☕ Casual`.
- **Instant `↩️ Undo`:** One-click rollback immediately restores the original document text on the page canvas.
- **Detached Floating Window (`popup-copilot.html`):** Optional floating OS window via `Office.context.ui.displayDialogAsync` that detaches and floats across multi-monitor setups.

### 2. Native Rich Word Formatting Engine (Zero Markdown Asterisks)
Standard AI integrations leak raw markdown syntax (`• **Client:**`, `**[Effective Date]**`) into the document. WordAI Pro includes a dedicated **Markdown-to-Word HTML Parser (`markdownToRichWordHtml`)**:
- `# Headings` $\rightarrow$ Styled native `Heading 1` (20pt bold, `#0f172a`, balanced margins).
- `## Headings` $\rightarrow$ Styled native `Heading 2` (14.5pt corporate accent, `#1e40af`).
- `### Headings` $\rightarrow$ Styled native `Heading 3` (12.5pt slate).
- `* **Key:** Description` $\rightarrow$ Native Word bulleted lists (`<ul><li>`) with genuine Word bold (`<strong>`).
- `1. ...` $\rightarrow$ Native Word numbered lists (`<ol><li>`).
- Multi-underscore signature lines (`By: __________________________`) are shielded with token masking (`@@USCORE...@@`), preventing accidental italicization.
- Markdown tables $\rightarrow$ Native Word table grids with shaded `#f1f5f9` header rows and borders.

### 3. Multi-Provider AI Suite (12 Providers — BYOK)
WordAI Pro connects directly to 12 frontier providers and local backends with live model fetching, customizable base URLs, and automatic failover chains:

```
                  ┌───────── WordAI Pro (taskpane.js) ─────────┐
                  │                                            │
        ┌─────────┴─────────┐                        ┌─────────┴─────────┐
   Cloud Frontier Providers                     Private & Local Backends
   • OpenAI (GPT-4o, o3-mini)                   • Ollama (Offline localhost)
   • Anthropic (Claude 3.5 Sonnet)              • LM Studio (Local OpenAI API)
   • Google Gemini (Gemini 2.0 Flash)           • vLLM / LocalAI
   • DeepSeek (V3, R1 Reasoning)                • Self-Hosted Azure OpenAI
   • Groq (Llama 3.3 70B @ 300 t/s)
   • Mistral AI (Mistral Large, Codestral)
   • Together AI (Qwen 2.5 72B, DeepSeek R1)
   • OpenRouter (Unified 200+ Models)
   • Perplexity AI (Sonar Online Search)
   • Cohere (Command R+)
```

### 4. Document Frameworks Studio (28 Pre-built Scaffolds)
Comprehensive document engineering studio across 5 critical domains:
- **🎓 Academic & Scientific (6):** *Research Proposal*, *Thesis (5 Chapters)*, *PRISMA Systematic Review*, *IMRaD Journal Article*, *Scientific Lab Report*, *Academic Case Study*.
- **💼 Business & Strategy (6):** *Executive Business Plan*, *Harvard Business Case Analysis*, *Project Charter & SOW*, *Standard Operating Procedure (SOP)*, *Quarterly Business Review (QBR)*, *White Paper*.
- **⚖️ Legal & Governance (3):** *Mutual Non-Disclosure Agreement (NDA)*, *Master Consulting Services Agreement (MSA)*, *Privacy Policy & GDPR/CCPA Notice*.
- **📢 Marketing & PR (3):** *AP-Style Press Release*, *Go-To-Market (GTM) Launch Plan*, *Executive Crisis Communication Plan*.
- **🎯 Career & Executive CV Studio (10):**
  1. 🏦 **Banking & Financial Services CV / Resume:** Credit risk underwriting, Basel III capital ratios, AML/KYC regulatory compliance, and core banking platforms (T24, Finacle).
  2. 👩‍🏫 **Teaching & Education Professional CV / Resume:** Differentiated learning (IEP/504), backward curriculum design (UbD), classroom leadership, EdTech LMS (Canvas, Google Classroom), and student testing gains.
  3. 🏥 **Healthcare, Nursing & Clinical Specialist CV:** Acute patient care, triage protocols, EHR systems (Epic, Cerner), JCAHO standards, ICU/ER telemetry, and HIPAA advocacy.
  4. 💻 **Modern Tech & Software Engineer CV:** Full-stack microservices, cloud infrastructure (AWS/GCP), distributed systems, CI/CD automation, and two-column technical skills matrix.
  5. 💼 **Executive Leadership & Management Resume:** P&L stewardship ($50M+), cross-functional operations, enterprise digital transformation, M&A integration, and Board governance.
  6. 📊 **Accounting, Audit & Corporate Finance Resume:** CPA/ACCA compliance, US GAAP/IFRS technical reporting, SOX 404 controls, external audit defense, and ERP (SAP S/4HANA, NetSuite).
  7. 📢 **Sales, Digital Marketing & Growth Resume:** High-quota sales pipeline generation ($10M+), CAC/LTV payback optimization, multi-channel SEO/SEM, and enterprise CRM closing.
  8. 🛠️ **Product & Agile Project Manager Resume:** Strategic product roadmaps, PRD authoring, RICE feature prioritization, Scrum/Kanban sprint velocity, and user funnel analytics.
  9. ⚖️ **Legal Counsel & Corporate Attorney CV:** High-stakes commercial contracts, M&A due diligence, GDPR/privacy counseling, risk management, and active state bar credentials.
  10. 🎓 **Academic Scholar & Faculty Curriculum Vitae:** Peer-reviewed publications in APA, NSF/ERC research grants ($1M+), doctoral mentorship, and university faculty appointments.

<p align="center">
  <img src="https://waseemkaluwal.github.io/wordai-pro/screenshots/screenshot-canvas-framework.png" alt="Document Framework Studio inserting structured Executive Business Plan directly onto Microsoft Word canvas" width="100%" />
</p>

**⚡ 3 Capabilities Per Framework:**
1. `Insert Skeleton ➕`: Instantly writes formatted titles, headings, bullet checklists, and signature tables into the Word body.
2. `🤖 AI Smart-Draft`: Context-aware generation tailoring the framework to your exact project parameters, depth, and tone.
3. `📋 Audit Doc`: Real-time compliance scanner verifying current document sections against the framework with a % score and 1-click missing section generation.

### 5. Pre-built Strategic Prompts Studio (48 Expert Prompts)
A dedicated repository of 48 battle-tested prompts categorized across 6 domains:
1. **🎓 Academic & Scientific Research (12):** Reviewer #2 Stress Test, Methodology Justification, Theoretical Framework, Limitations & Delimitations, Statistical Findings Narrative (APA 7th), Qualitative Thematic Analysis.
2. **💼 Business & Executive Strategy (10):** McKinsey Pyramid Principle (Answer First), SWOT & Moat Analysis, Board Memo, VC Pitch Deck Flow, Crisis Holding Statement.
3. **⚖️ Legal & Contract Analysis (6):** Ambiguity Scanner, Plain-English Translator, GDPR Compliance Audit, NDA Carve-Out Review, IP Assignment Verification.
4. **📝 Editing & Tone Polish (8):** Harvard/Oxford Voice Polish, Jargon to 8th-Grade Clarity, Passive-to-Active Converter, Executive Brevity 40% Compression.
5. **🎯 Career & Job Hunting (6):** Google XYZ Resume Bullet Formula, ATS Keyword Alignment, STAR Interview Method, Salary Counter-Offer Script.
6. **📣 Marketing & Copywriting (6):** PAS Framework, AIDA Copy, Viral Social Thread, Objection Handling Matrix.

### 6. Visual Redline Diff & Side-by-Side Evaluation
- **3-Way Output Switcher:** Toggle between **Preview**, **Redline Diff**, and **Side-by-Side** before accepting changes into Word.
- **LCS Redline Diff Engine:** Highlights deleted phrases in red strikethrough and new content in emerald green.
- **Side-by-Side Evaluator:** Two-column synchronized card comparing original text vs. AI revision with independent copy/insert controls.

### 7. The 6 Killer Features Suite (New in v1.0.1)
WordAI Pro v1.0.1 introduces six game-changing superpowers designed to surpass proprietary commercial tools:
1. **💬 Chat With Your Document (Full Document Q&A):** Conversational assistant connected directly to the open Word document's full body text. Features real-time word count sync and 6 one-click deep-analysis chips (*Summarize*, *Risks & Gaps*, *Dates & Deadlines*, *Action Items*, *Exec Brief*, *Critical Review*). Every AI response includes a 1-click **"➕ Insert to Word"** button.
2. **🪄 AI Humanizer & Anti-Robotic Polish Engine:** Targets and eliminates telltale AI transition clichés (*"delve into"*, *"testament to"*, *"rich tapestry"*, *"in conclusion"*, *"moreover"*). Optimizes sentence burstiness and presents an interactive **Originality Rating Card (98% Human Score)**.
3. **🎓 1-Click Academic Citation & Auto-Bibliography Builder:** Formats citations in **APA 7th**, **MLA 9th**, **Chicago**, **Harvard**, and **IEEE**. Features automated **"Append to Bibliography"** which creates or locates the `# References` heading at the very end of the Word document and appends references with hanging indents.
4. **🧬 Personal Writing DNA & Brand Voice Studio:** Analyzes writing samples to profile formality, sentence rhythm, and dominant authorial persona, saving a persistent voice profile that shapes all subsequent outputs.
5. **📊 Text-to-Native Word Table Generator:** Automatically extracts unstructured text, metrics, or comparison data and inserts clean, styled Word tables with colored headers (`#4f46e5`) and alternating shaded rows (`#f8fafc`).
6. **🌐 Format-Preserving Multi-Language Translator:** High-fidelity translation supporting 30+ languages that strictly retains all HTML and markdown formatting (bolding, italics, headings, bullet hierarchies, and tables).
7. **🔍 Resilient Academic Search Engine:** Multi-tiered search querying **OpenAlex** (250M+ scholarly works) and **Crossref DOI Registry** with AI synthesis fallback, eliminating 429 rate limit and fetch errors.

---

## 🚀 Complete 36-Feature Capabilities Matrix

| # | Feature | Category | Description |
|---|---|---|---|
| **1** | **Plagiarism & Originality Highlighter** | High Priority | Flags clichéd/unoriginal phrasing with originality % and fresh academic rewrites. |
| **2** | **AI Citation Finder** | High Priority | Suggests 3–5 real papers/books with direct links to Google Scholar, PubMed, and JSTOR. |
| **3** | **Document Frameworks Studio** | High Priority | 28 full frameworks with skeleton insertion, smart drafting, and compliance audits. |
| **4** | **Word Count Target Tracker** | High Priority | Monospace live word count, target goal progress bar, and percentage completion. |
| **5** | **Reading Level Analyzer** | High Priority | Flesch Reading Ease, Flesch-Kincaid Grade Level, and estimated reading duration. |
| **6** | **Research Gap Finder** | Research Copilot | Analyzes abstracts to identify unexamined populations and methodological gaps. |
| **7** | **Thesis Statement Generator** | Research Copilot | Generates 5 distinct, debate-ready thesis formats (Analytical, Argumentative, Policy, etc.). |
| **8** | **Counter-Argument Builder** | Research Copilot | Formulates the 3 strongest opposing views with counter-evidence and rebuttal blueprints. |
| **9** | **Interview Questions Protocol** | Research Copilot | Qualitative interview protocol with rapport warm-ups, core probes, and debriefing. |
| **10** | **Survey & Questionnaire Builder** | Research Copilot | Formulates balanced 5-point Likert scale surveys free of acquiescence bias. |
| **11** | **Sentence Rewriter (3 Versions)** | Writing Assistant | Generates Concise, Formal, and Creative versions side-by-side with 1-click replace. |
| **12** | **Transition Sentence Architect** | Writing Assistant | Analyzes consecutive paragraphs and builds smooth conceptual bridge transitions. |
| **13** | **Paragraph Expander with Evidence**| Writing Assistant | Expands conceptual density with explicit `[ADD EVIDENCE HERE]` empirical slots. |
| **14** | **Jargon & Buzzword Detector** | Writing Assistant | Flags bureaucratic jargon, generates translation tables, and simplifies language. |
| **15** | **Document Consistency Checker** | Writing Assistant | Validates acronym usage, straight vs. curly quotes, and mixed numbering schemes. |
| **16** | **Section-by-Section Summarizer** | Doc Intelligence | Traverses native Word headings (H1/H2/H3) to produce modular section digests. |
| **17** | **Key Terms & Glossary Extractor** | Doc Intelligence | Compiles alphabetical index of 15–20 technical terms with one-sentence definitions. |
| **18** | **Action Items & Deliverables** | Doc Intelligence | Extracts structured tables with tasks, assignees, deliverables, and urgency levels. |
| **19** | **Document Comparison & Evaluator** | Doc Intelligence | Evaluates two passages, scores readability, picks a winner, and synthesizes a hybrid. |
| **20** | **Sentiment & Emotional Tone** | Doc Intelligence | Scores sentiment (-100 to +100) with section-by-section emotional breakdown. |
| **21** | **Email Mode Studio** | Dedicated Modes | 8 professional email templates + 4-stance interactive reply drafter. |
| **22** | **Job Application & Career Mode** | Dedicated Modes | Google XYZ resume bullets, tailored cover letters, and STAR interview answers. |
| **23** | **Social Media Repurposing Mode** | Dedicated Modes | Converts document excerpts into viral LinkedIn posts, X threads, and newsletters. |
| **24** | **Presentation & Slide Mode** | Dedicated Modes | Generates 8-slide decks, investor pitches, and executive briefs with speaker notes. |
| **25** | **Multi-Language Side-by-Side** | Dedicated Modes | 2-column comparative translation (English, Urdu, Arabic, Spanish, French, German). |
| **26** | **Prompt History Drawer** | Power User | Off-canvas drawer caching the last 20 prompts with 1-click re-execution. |
| **27** | **Custom Prompt Presets** | Power User | Create, name, save, and manage reusable custom system instructions. |
| **28** | **Auto-Save Session Log** | Power User | Persistent drawer storing the last 30 AI generations with 1-click reload and insert. |
| **29** | **Global Keyboard Shortcuts** | Power User | Complete keyboard control (`Ctrl+Shift+E`, `R`, `S`, `I`, `G`, `D`, `H`). |
| **30** | **Obsidian Dark Mode** | Power User | Midnight Obsidian theme (`#0b0d14`) with persistent theme state. |
| **31** | **Chat With Document (Full Q&A)** | Killer Features | Conversational document analysis with real-time word count sync and 1-click canvas insertion. |
| **32** | **AI Humanizer & Anti-Robotic Polish** | Killer Features | Eliminates AI clichés, optimizes sentence burstiness, and provides an interactive 98% Human Scorecard. |
| **33** | **1-Click Auto-Bibliography Builder** | Killer Features | Formats APA, MLA, Harvard, IEEE citations and appends them to `# References` at the document end. |
| **34** | **Personal Writing DNA Studio** | Killer Features | Extracts formality, sentence rhythm, and dominant author persona to match your exact writing style. |
| **35** | **Text-to-Native Word Table** | Killer Features | Extracts data from paragraphs/lists and inserts styled Word tables with colored headers and alternating shading. |
| **36** | **Format-Preserving Translator** | Killer Features | Translates into 30+ languages with strict preservation of all bold, italic, header, and table tags. |

---

## ⌨️ Global Keyboard Shortcuts

| Shortcut | Action | Scope |
| :--- | :--- | :--- |
| **`Ctrl + Shift + E`** | **Open In-Place Floating Copilot Popup** | Word Document & Taskpane |
| **`Ctrl + Shift + R`** | **Quick Rewrite Selection** | Word Document |
| **`Ctrl + Shift + S`** | **Quick Summarize Selection** | Word Document |
| **`Ctrl + Shift + I`** | **Quick Improve Selection** | Word Document |
| **`Ctrl + Shift + G`** | **Trigger Main Prompt Generation** | Taskpane |
| **`Ctrl + Shift + D`** | **Toggle Dark / Light Mode** | Global |
| **`Ctrl + Shift + H`** | **Toggle History & Presets Drawer** | Global |
| **`Escape`** | **Close Floating Popup Modal / Drawer** | Global |

---

## 💻 Installation & Setup Guide

### Method 1: Word Online (Free — 60 Seconds)
*No developer tools, installations, or subscriptions required. Works in Chrome, Edge, Safari, and Firefox.*

1. Navigate to **[office.com](https://www.office.com)** and sign in with any free or school/work Microsoft account.
2. Open any existing Word document or create a **New Blank Document**.
3. In the top ribbon, select **Insert** $\rightarrow$ **Add-ins** $\rightarrow$ **My Add-ins**.
4. Click **Upload My Add-in** in the top-right corner of the dialog.
5. Browse and select [`manifest.xml`](manifest.xml) from this repository.
6. The **WordAI Pro** icon will appear on your ribbon. Click it to launch the taskpane.
7. Click **⚙️ Settings** in the add-in header $\rightarrow$ select your provider (e.g. Gemini, OpenAI, Groq, or Ollama) $\rightarrow$ enter your key $\rightarrow$ click **Save Settings**.

<p align="center">
  <img src="https://waseemkaluwal.github.io/wordai-pro/screenshots/screenshot-sideload.png" alt="Sideloading WordAI Pro manifest.xml directly into Microsoft Word" width="100%" />
</p>

---

### Method 2: Word Desktop (Windows / macOS)
*Prerequisites: [Node.js](https://nodejs.org/) (v16+) and Microsoft Word (Office 2016+, 2019+, 2021+, or Microsoft 365).*

```bash
# 1. Clone the repository
git clone https://github.com/waseemkaluwal/wordai-pro.git
cd wordai-pro

# 2. Install dependencies
npm install

# 3. Trust self-signed developer certificates
npx office-addin-dev-certs install

# 4. Launch dev server and sideload into Word
npm start
```

---

### Deploying to GitHub Pages (Single Root Folder)
WordAI Pro is engineered with **zero subfolder dependencies**—all logos, HTML, CSS, and JS files live in the root directory:

1. Fork or push this repository to your GitHub account (`https://github.com/YOUR_USERNAME/wordai-pro`).
2. Navigate to repository **Settings** $\rightarrow$ **Pages**.
3. Under **Build and deployment**, set Source to **Deploy from a branch** $\rightarrow$ select `main` branch $\rightarrow$ folder `/ (root)` $\rightarrow$ click **Save**.
4. Update [`manifest.xml`](manifest.xml) to replace `waseemkaluwal` with your GitHub username.
5. Upload the updated `manifest.xml` to Word!

---

## 🔒 Privacy Architecture & Security

- **100% Client-Side Architecture (BYOK):** WordAI Pro has **no middleman proxy, no cloud database, and zero analytics servers**. All API requests are dispatched directly from your local browser/Word sandbox to your chosen AI provider using standard HTTPS.
- **Zero Document Retention:** Your documents are never logged, cached, or used for model training by WordAI Pro.
- **Local Air-Gapped AI (Ollama / LM Studio):** For confidential, medical, legal, or sensitive government documents, switch the provider to **Ollama** (`http://localhost:11434/v1`). Your document text never leaves your physical workstation.
- **Credential Protection:** API keys are stored exclusively in your browser's encrypted local storage (`localStorage`) and are never transmitted anywhere except the provider's official endpoint.

---

## ❓ Frequently Asked Questions (FAQ)

<details>
<summary><strong>1. Does WordAI Pro write directly onto the Word document page canvas?</strong></summary>

**Yes, 100%.** WordAI Pro uses Microsoft Word's native JavaScript API (`Word.Range.insertHtml`, `Word.Body.insertHtml`). When you click **Replace ⚡**, **Insert ✅**, or **Continue Writing**, the AI writes and replaces text directly inside your paragraphs on the page canvas in real time.
</details>

<details>
<summary><strong>2. Can I use WordAI Pro completely free without paid API keys?</strong></summary>

**Yes.** You have two free options:
1. **Google Gemini:** Google AI Studio offers a free tier (free API key with 15 requests per minute).
2. **Ollama (100% Free & Offline):** Install [Ollama](https://ollama.com/) on your computer, run `ollama run llama3.2` or `deepseek-r1`, and select **Ollama** in WordAI Pro settings. It runs 100% locally with zero costs and zero API keys.
</details>

<details>
<summary><strong>3. Why does WordAI Pro not show raw markdown asterisks (`**`) in my Word document?</strong></summary>

WordAI Pro features a built-in **Rich Word HTML Engine (`markdownToRichWordHtml`)**. It converts markdown into Word-compliant semantic HTML before insertion. Word natively parses this into real bold runs (`<strong>`), authentic Word heading styles (`Heading 1`, `Heading 2`), native Word bullet lists, and genuine formatted tables.
</details>

<details>
<summary><strong>4. Which Microsoft Word versions are supported?</strong></summary>

- Word on Windows (Office 2013, 2016, 2019, 2021, and Microsoft 365)
- Word on Mac (Word 2016, 2019, 2021, and Microsoft 365)
- Word on the Web (office.com via Chrome, Edge, Firefox, Safari)
- Word on iPad
</details>

<details>
<summary><strong>5. How do I open the In-Place Copilot Popup?</strong></summary>

Highlight any sentence in your Word document and press **`Ctrl + Shift + E`**, or click the **`⚡ Popup`** button located right beneath the header in the sidebar.
</details>

---

## 📁 Repository File Structure

```text
WordAI Pro/
├── screenshots/       # High-resolution UI showcase & sideloading screenshots
├── manifest.xml       # Office Add-in manifest configuration (Validated across all platforms)
├── package.json       # Dependencies, build scripts & dev cert utilities
├── README.md          # Comprehensive documentation, API guide & user manual
├── commands.html      # Office runtime command host
├── generate-icons.js  # Zero-dependency PNG icon generator
├── ai-provider.js     # Multi-provider AI communication engine (12 providers)
├── frameworks-data.js # 28 Pre-built Document Frameworks with skeletons & audit specs
├── prompts-data.js    # 48 Curated Pre-built Strategic Prompts across 6 domains
├── popup-copilot.html # Detached floating in-place copilot dialog
├── taskpane.html      # Add-in UI structure (Modes, Framework Studio, History Drawer)
├── taskpane.css       # Complete responsive design system, obsidian dark mode & animations
├── taskpane.js        # Core logic, Word.run bridge, NLP tools, shortcuts & storage
├── icon-16.png        # Root add-in icon (16x16)
├── icon-32.png        # Root add-in icon (32x32)
└── icon-80.png        # Root add-in icon (80x80)
```

---

## 🤝 Contributing & License

Contributions, feature requests, and bug reports are warmly welcomed!
- **Issues & Discussions:** [GitHub Issues](https://github.com/waseemkaluwal/wordai-pro/issues)
- **License:** Free for personal, academic, and internal organizational use under the **WordAI Pro Non-Commercial License** (commercial resale or rebranding prohibited). See [`LICENSE`](LICENSE) for details.

---

<p align="center">
  <b>Built with ❤️ for students, academic researchers, and professional writers worldwide.</b><br>
  <sub>Microsoft Word is a registered trademark of Microsoft Corporation. WordAI Pro is an independent open-source project.</sub>
</p>
