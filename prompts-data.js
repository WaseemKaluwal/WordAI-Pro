// prompts-data.js — WordAI Pro Pre-built Prompts Repository (48 Expert Prompts)

const PREBUILT_PROMPTS = [
  // ─── 🎓 ACADEMIC & SCIENTIFIC RESEARCH (12) ─────────────────────────────
  {
    id: 'lit-review-matrix',
    category: 'academic',
    categoryLabel: 'Research',
    title: 'Evidence Matrix & Research Gap Synthesis',
    desc: 'Analyzes texts/abstracts to construct a comparative evidence matrix and pinpoint unaddressed gaps.',
    prompt: 'Analyze this literature review text or collection of abstracts. Construct a Markdown Evidence Matrix comparing: Author/Year, Theoretical Stance, Methodology, Key Findings, and identify at least 3 critical unaddressed research gaps that warrant further study.'
  },
  {
    id: 'methodology-defense',
    category: 'academic',
    categoryLabel: 'Research',
    title: 'Methodology Justification & Defense',
    desc: 'Provides a formal epistemological justification explaining why this research design is optimal.',
    prompt: 'Critique and justify the research methodology in this draft. Provide a formal epistemological defense explaining why this design (qualitative/quantitative/mixed) is superior to alternative paradigms for answering the primary research questions.'
  },
  {
    id: 'reviewer-2-critique',
    category: 'academic',
    categoryLabel: 'Research',
    title: 'Reviewer #2 Harsh Peer Review',
    desc: 'Performs a rigorous peer-review stress test identifying unsupported claims and confounding variables.',
    prompt: 'Act as Reviewer #2 for an elite peer-reviewed academic journal. Scrutinize this excerpt with rigorous academic skepticism: identify unsupported assertions, methodological vulnerabilities, potential confounding variables, and recommend 3 concrete structural remedies.'
  },
  {
    id: 'theoretical-framework',
    category: 'academic',
    categoryLabel: 'Research',
    title: 'Theoretical Framework Formulator',
    desc: 'Constructs a conceptual model connecting independent, mediating, and dependent variables.',
    prompt: 'Based on the following research problem and variables, formulate a coherent Theoretical Framework. Propose 2-3 established foundational theories (with seminal citations) and construct a logical conceptual model connecting the independent, mediating, and dependent variables.'
  },
  {
    id: 'limits-delimitations',
    category: 'academic',
    categoryLabel: 'Research',
    title: 'Limitations & Delimitations Section',
    desc: 'Drafts an honest, rigorous Limitations and Delimitations section without undermining core contributions.',
    prompt: 'Formulate an academically honest Limitations and Delimitations section for this research draft. Categorize internal validity limits, external generalizability bounds, and methodological trade-offs without undermining the study\'s core contributions.'
  },
  {
    id: 'apa7-findings-interpreter',
    category: 'academic',
    categoryLabel: 'Research',
    title: 'Statistical Findings Narrative (APA 7th)',
    desc: 'Converts quantitative data/tables into an APA 7th narrative reporting significance and effect sizes.',
    prompt: 'Convert these quantitative findings / data tables into a formal APA 7th style narrative report. Report effect sizes, statistical significance (p-values), confidence intervals, and interpret what the numbers mean in relation to the primary hypotheses.'
  },
  {
    id: 'thematic-coding',
    category: 'academic',
    categoryLabel: 'Research',
    title: 'Qualitative Thematic Analysis & Coding',
    desc: 'Extracts primary themes, sub-themes, and supporting verbatim quotes from qualitative text.',
    prompt: 'Perform an inductive thematic analysis on this qualitative interview/observation text. Extract primary thematic codes, secondary sub-themes, and provide supporting verbatim quotes with analytical commentary.'
  },
  {
    id: 'dialectic-counter-narrative',
    category: 'academic',
    categoryLabel: 'Research',
    title: 'Academic Counter-Narrative & Rebuttal',
    desc: 'Constructs strong opposing arguments and crafts sophisticated dialectical rebuttals.',
    prompt: 'Construct the strongest possible academic counter-arguments against the thesis presented in this text. Then, formulate a sophisticated dialectical rebuttal for each counter-argument that strengthens the author\'s original position.'
  },
  {
    id: 'grant-broader-impacts',
    category: 'academic',
    categoryLabel: 'Research',
    title: 'Grant Proposal Broader Impacts Statement',
    desc: 'Drafts a compelling Intellectual Merit & Broader Impacts section for national funding proposals.',
    prompt: 'Transform this research summary into a compelling National Grant Proposal "Intellectual Merit and Broader Impacts" statement. Emphasize societal benefit, scientific novelty, cross-disciplinary scalability, and feasibility.'
  },
  {
    id: 'annotated-bib-entry',
    category: 'academic',
    categoryLabel: 'Research',
    title: 'Annotated Bibliography Synthesizer',
    desc: 'Synthesizes this source into an evaluation of thesis, methodology, and field contribution.',
    prompt: 'Synthesize this source into a formal 150-word Annotated Bibliography entry: summarize the central thesis, evaluate methodological rigor, and assess its unique contribution to the broader field.'
  },
  {
    id: 'viva-oral-defense-prep',
    category: 'academic',
    categoryLabel: 'Research',
    title: 'Oral Defense / Viva Voce Committee Q&A',
    desc: 'Generates 7 probing defense questions and strategic answering frameworks for your thesis committee.',
    prompt: 'Generate 7 probing, difficult defense questions that an academic examination committee would ask regarding this draft, along with strategic framework bullet points for answering each effectively.'
  },
  {
    id: 'prisma-search-protocol',
    category: 'academic',
    categoryLabel: 'Research',
    title: 'PRISMA Systematic Review Protocol',
    desc: 'Builds Boolean search strings and explicit inclusion/exclusion criteria for systematic reviews.',
    prompt: 'Construct a formal PRISMA-compliant search protocol for this research question: formulate Boolean search strings (AND/OR/NOT) for PubMed, Scopus, and Web of Science, plus explicit Inclusion and Exclusion criteria.'
  },

  // ─── 💼 BUSINESS & EXECUTIVE STRATEGY (10) ──────────────────────────────
  {
    id: 'pyramid-exec-summary',
    category: 'business',
    categoryLabel: 'Business',
    title: 'McKinsey Pyramid Principle Summary',
    desc: 'Structures summary with Answer First (Governing Thought) followed by 3 data-backed arguments.',
    prompt: 'Rewrite this document using the McKinsey Pyramid Principle: lead with the single Overarching Governing Thought (Answer First), followed by 3 structured supporting arguments backed by key data points.'
  },
  {
    id: 'swot-moat-analysis',
    category: 'business',
    categoryLabel: 'Business',
    title: 'SWOT & Defensible Moat Analysis',
    desc: 'Evaluates strengths, weaknesses, opportunities, threats, and defensible competitive moats.',
    prompt: 'Conduct a strategic SWOT and Porter\'s 5 Forces evaluation based on this business context. Specifically identify the company\'s defensible competitive moats and high-probability strategic risks.'
  },
  {
    id: 'board-briefing-memo',
    category: 'business',
    categoryLabel: 'Business',
    title: 'Board of Directors Briefing Memo',
    desc: 'Condenses operational reports into a high-impact 1-page memo for board members and executives.',
    prompt: 'Condense this operational report into a crisp 1-page Board of Directors Executive Briefing: Context, Strategic Implications, Financial/Resource Impact, and Decisions Required.'
  },
  {
    id: 'business-case-roi',
    category: 'business',
    categoryLabel: 'Business',
    title: 'Business Case & ROI Justification',
    desc: 'Formulates a business case with cost-benefit analysis, payback timeline, and strategic ROI.',
    prompt: 'Develop a formal business case for this project initiative: articulate the problem statement, proposed solution, projected ROI/payback period, cost-benefit analysis, and implementation roadmap.'
  },
  {
    id: 'pitch-deck-narrative',
    category: 'business',
    categoryLabel: 'Business',
    title: 'VC Investor Pitch Deck Narrative',
    desc: 'Structures a 10-slide VC pitch flow covering Problem, Solution, Market (TAM), Model, and The Ask.',
    prompt: 'Structure this business idea into a 10-slide venture capital pitch deck narrative: Problem, Solution, Market Size (TAM/SAM/SOM), Product, Traction, Business Model, Competition, Team, Financials, The Ask.'
  },
  {
    id: 'post-mortem-retro',
    category: 'business',
    categoryLabel: 'Business',
    title: 'Project Post-Mortem & Retrospective',
    desc: 'Synthesizes project deliverables, root-cause bottlenecks, and actionable sprint learnings.',
    prompt: 'Analyze this project delivery summary: synthesize What Went Well, What Missed Expectations, Root Cause Analysis for bottlenecks, and 5 Actionable Lessons Learned for future sprints.'
  },
  {
    id: 'crisis-comms-plan',
    category: 'business',
    categoryLabel: 'Business',
    title: 'Crisis Communications & Holding Statement',
    desc: 'Drafts emergency crisis communication plans and transparent public holding statements.',
    prompt: 'Draft an executive crisis communication plan and public holding statement for this incident: transparently address impacted parties, outline immediate remediations, and project calm leadership.'
  },
  {
    id: 'kotter-change-mgmt',
    category: 'business',
    categoryLabel: 'Business',
    title: 'Kotter Change Management Briefing',
    desc: 'Frames organizational change using Kotter\'s 8-step transformation and alignment model.',
    prompt: 'Formulate an organizational change management briefing for this transition using Kotter\'s 8-step framework: creating urgency, guiding coalition, communication vision, and cultural anchoring.'
  },
  {
    id: 'iso-sop-guide',
    category: 'business',
    categoryLabel: 'Business',
    title: 'ISO Standard Operating Procedure (SOP)',
    desc: 'Converts workflows into formal step-by-step procedures with checkpoints and exceptions.',
    prompt: 'Convert this procedural description into a formal ISO-style Standard Operating Procedure (SOP): Scope, Responsibilities, Step-by-Step Workflow, Quality Checkpoints, and Troubleshooting exceptions.'
  },
  {
    id: 'qbr-synthesis',
    category: 'business',
    categoryLabel: 'Business',
    title: 'Quarterly Business Review (QBR)',
    desc: 'Synthesizes quarterly updates into OKR achievements, financial metrics, and next-quarter targets.',
    prompt: 'Synthesize these departmental updates into a Quarterly Business Review (QBR) presentation draft: Key OKR achievements, revenue/efficiency metrics, headwinds encountered, and Q+1 strategic priorities.'
  },

  // ─── ⚖️ LEGAL & CONTRACT ANALYSIS (6) ───────────────────────────────────
  {
    id: 'contract-risk-scanner',
    category: 'legal',
    categoryLabel: 'Legal',
    title: 'Contract Ambiguity & Risk Scanner',
    desc: 'Identifies unilateral clauses, indemnification traps, and suggests balanced protective redlines.',
    prompt: 'Review this contract clause or agreement: highlight ambiguous terms, unilateral obligations, indemnification pitfalls, and hidden liability traps. Suggest redline revision language for balanced protection.'
  },
  {
    id: 'plain-english-contract',
    category: 'legal',
    categoryLabel: 'Legal',
    title: 'Plain-English Legal Translator',
    desc: 'Translates dense legal terminology into clear executive terms with obligations and deadlines.',
    prompt: 'Translate this dense legalese agreement into clear, plain-English terms that a non-lawyer business executive can understand, highlighting obligations, deadlines, and termination triggers.'
  },
  {
    id: 'gdpr-compliance-gap',
    category: 'legal',
    categoryLabel: 'Legal',
    title: 'GDPR / CCPA Privacy Compliance Audit',
    desc: 'Audits data processing and privacy terms against European GDPR and California CCPA standards.',
    prompt: 'Audit this privacy policy or data processing terms against GDPR/CCPA requirements: check lawful basis, data retention disclosures, third-party transfers, and data subject rights enforcement.'
  },
  {
    id: 'nda-redline-audit',
    category: 'legal',
    categoryLabel: 'Legal',
    title: 'NDA Redline & Carve-Out Review',
    desc: 'Audits non-disclosure terms, definition of confidential information, and survival clauses.',
    prompt: 'Review this NDA excerpt: check definition of confidential information, carve-outs (standard exclusions), duration of confidentiality, and non-solicitation restrictions. Recommend pro-recipient redlines.'
  },
  {
    id: 'sla-remedies-eval',
    category: 'legal',
    categoryLabel: 'Legal',
    title: 'SLA Uptime & Service Credit Audit',
    desc: 'Evaluates uptime percentages, measurement windows, and credit remedies for fairness.',
    prompt: 'Analyze these SLA terms: evaluate uptime percentages, measurement windows, service credit calculations, and dispute resolution mechanisms for operational fairness.'
  },
  {
    id: 'ip-work-for-hire',
    category: 'legal',
    categoryLabel: 'Legal',
    title: 'IP Assignment & Work-Made-For-Hire Audit',
    desc: 'Ensures comprehensive assignment of copyrights, patents, and waiver of moral rights.',
    prompt: 'Examine this IP clause: ensure absolute assignment of copyrights, patents, and trade secrets, and confirm waiver of moral rights with comprehensive survival clauses.'
  },

  // ─── 📝 EDITING, TONE & REWRITING (8) ───────────────────────────────────
  {
    id: 'academic-voice-polish',
    category: 'editing',
    categoryLabel: 'Editing',
    title: 'Harvard/Oxford Academic Voice Polish',
    desc: 'Polishes text into elite academic scholarly tone with heightened lexical precision.',
    prompt: 'Polish this text to conform to elite academic standards: remove colloquialisms, enhance lexical precision, ensure scholarly objectivity, and eliminate first-person subjectivity.'
  },
  {
    id: 'jargon-clarity-rewrite',
    category: 'editing',
    categoryLabel: 'Editing',
    title: 'Jargon to 8th-Grade Clarity Rewrite',
    desc: 'Rewrites dense technical jargon into crystal clear prose without sacrificing factual accuracy.',
    prompt: 'Translate this overly complex, jargon-heavy text into clear, engaging prose understandable by an 8th-grade reader without sacrificing technical accuracy.'
  },
  {
    id: 'passive-to-active-verbs',
    category: 'editing',
    categoryLabel: 'Editing',
    title: 'Passive-to-Active Dynamic Verbs',
    desc: 'Replaces sluggish passive sentences with energetic active verbs that convey momentum.',
    prompt: 'Transform all passive voice constructions in this text into dynamic, active voice verbs that convey agency, urgency, and direct impact.'
  },
  {
    id: 'sentence-rhythm-cadence',
    category: 'editing',
    categoryLabel: 'Editing',
    title: 'Rhetorical Rhythm & Cadence Variance',
    desc: 'Mixes staccato impact sentences with flowing clauses for irresistible reading rhythm.',
    prompt: 'Rewrite this passage to vary sentence length dramatically (combining short staccato sentences with rhythmic complex clauses) to maximize reader engagement and cadence.'
  },
  {
    id: 'buzzword-exterminator',
    category: 'editing',
    categoryLabel: 'Editing',
    title: 'Corporate Buzzword Exterminator',
    desc: 'Strips out corporate jargon ("synergy", "paradigm") and substitutes concrete evidence.',
    prompt: 'Eliminate all corporate buzzwords ("synergy", "circle back", "move the needle", "paradigm shift") and clichés from this text, replacing them with concrete, substantive facts.'
  },
  {
    id: 'executive-compression',
    category: 'editing',
    categoryLabel: 'Editing',
    title: 'Executive Brevity 40% Compression',
    desc: 'Compresses text by 40% while preserving 100% of the core insights and takeaways.',
    prompt: 'Compress this text by 40% while preserving 100% of the key takeaways. Make every sentence direct, punchy, and executive-ready.'
  },
  {
    id: 'transition-architect',
    category: 'editing',
    categoryLabel: 'Editing',
    title: 'Cohesive Paragraph Transition Architect',
    desc: 'Constructs smooth narrative bridges connecting disconnected paragraphs seamlessly.',
    prompt: 'Construct smooth, logical bridge and transition sentences between these paragraphs so the document reads like a cohesive, uninterrupted narrative.'
  },
  {
    id: 'metaphor-explainer',
    category: 'editing',
    categoryLabel: 'Editing',
    title: 'Analogy & Metaphor Conceptual Clarifier',
    desc: 'Creates 2 intuitive real-world metaphors to explain abstract concepts clearly.',
    prompt: 'Create 2 vivid real-world analogies or metaphors to explain the complex concept described in this passage to non-specialist readers.'
  },

  // ─── 🎯 CAREER & JOB HUNTING (6) ────────────────────────────────────────
  {
    id: 'google-xyz-resume',
    category: 'career',
    categoryLabel: 'Career',
    title: 'Google XYZ Resume Bullet Formula',
    desc: 'Rewrites accomplishments into "Accomplished [X] measured by [Y] by doing [Z]".',
    prompt: 'Transform this job accomplishment into the Google XYZ format: "Accomplished [X] as measured by [Y], by doing [Z]". Include high-impact action verbs and quantified metrics.'
  },
  {
    id: 'job-keyword-matcher',
    category: 'career',
    categoryLabel: 'Career',
    title: 'ATS Keyword Alignment Scanner',
    desc: 'Compares resume bullets against target job requirements to close critical keyword gaps.',
    prompt: 'Compare this resume section against the provided target job description. Identify missing ATS keywords, critical skills gaps, and rewrite 3 bullets to mirror the required competencies.'
  },
  {
    id: 'star-interview-response',
    category: 'career',
    categoryLabel: 'Career',
    title: 'STAR Method Behavioral Response',
    desc: 'Formats experiences into Situation (15%), Task (15%), Action (50%), and Result (20%).',
    prompt: 'Structure this work experience into a powerful STAR method interview answer: Situation (15%), Task (15%), Action (50%), and Result (20% with quantitative business impact).'
  },
  {
    id: 'salary-counter-script',
    category: 'career',
    categoryLabel: 'Career',
    title: 'Salary Negotiation Counter-Offer Script',
    desc: 'Drafts a confident, professional email negotiating a 10-15% increase or equity bump.',
    prompt: 'Draft a professional, confident counter-offer negotiation email for this job offer: express gratitude, cite market benchmarks, and justify a 10-15% increase or equity adjustment.'
  },
  {
    id: 'linkedin-thought-leader',
    category: 'career',
    categoryLabel: 'Career',
    title: 'LinkedIn Thought Leadership Post',
    desc: 'Transforms document insights into a high-engagement LinkedIn post with hook and CTA.',
    prompt: 'Convert this document insight into an engaging LinkedIn article: strong hook, 3 digestible bulleted takeaways, personal perspective, and an open conversation-starting question.'
  },
  {
    id: 'inmail-networking-pitch',
    category: 'career',
    categoryLabel: 'Career',
    title: '75-Word Warm InMail Networking Pitch',
    desc: 'Crafts a low-friction 75-word outreach pitch to senior executives for coffee chats.',
    prompt: 'Write a concise 75-word warm networking message to a senior executive: cite a shared interest, compliment their recent work, and propose a low-friction 10-minute virtual coffee.'
  },

  // ─── 📣 MARKETING & COPYWRITING (6) ─────────────────────────────────────
  {
    id: 'pas-copy-formula',
    category: 'marketing',
    categoryLabel: 'Marketing',
    title: 'PAS (Problem-Agitate-Solution) Copy',
    desc: 'Structures persuasive copy by framing the Problem, Agitating friction, and offering Solution.',
    prompt: 'Rewrite this product pitch using the PAS framework: clearly define the customer\'s burning Problem, Agitate the cost of inaction, and present the Solution as the obvious answer.'
  },
  {
    id: 'aida-sales-message',
    category: 'marketing',
    categoryLabel: 'Marketing',
    title: 'AIDA (Attention-Interest-Desire-Action)',
    desc: 'Crafts high-converting sales messaging flowing from Attention to Action.',
    prompt: 'Format this offer into an AIDA sales message: grab Attention with a bold hook, build Interest with unique mechanisms, create Desire with transformation proof, and drive Action with a clear CTA.'
  },
  {
    id: 'viral-thread-builder',
    category: 'marketing',
    categoryLabel: 'Marketing',
    title: '7-Post Viral Social Thread (X/Threads)',
    desc: 'Transforms documents into an educational viral thread with hooks and numbered insights.',
    prompt: 'Convert this document into a 7-post educational viral social thread: punchy opening hook, curated insights, counter-intuitive facts, and a concluding bookmark/share call.'
  },
  {
    id: 'objection-counter-matrix',
    category: 'marketing',
    categoryLabel: 'Marketing',
    title: 'Customer Objection Buster Matrix',
    desc: 'Identifies top 5 buyer hesitations and drafts persuasive, non-defensive responses.',
    prompt: 'Identify the top 5 objections a prospective buyer would have to this proposal (price, timing, trust, complexity, alternatives) and draft persuasive, respectful rebuttals for each.'
  },
  {
    id: 'seo-pillar-faq',
    category: 'marketing',
    categoryLabel: 'Marketing',
    title: 'SEO Pillar Outline & FAQ Schema',
    desc: 'Creates an SEO-optimized heading structure with schema-ready FAQ questions and answers.',
    prompt: 'Generate an SEO-optimized pillar page outline for this topic: H1, primary H2s targeting search intent, secondary H3s, and 5 schema-ready FAQ questions with concise answers.'
  },
  {
    id: 'newsletter-story-hook',
    category: 'marketing',
    categoryLabel: 'Marketing',
    title: 'Newsletter Storytelling Hook & Intro',
    desc: 'Converts dry document data into a compelling personal story that hooks readers immediately.',
    prompt: 'Turn this dry report topic into a captivating personal storytelling newsletter opening that hooks the reader within the first 3 seconds before transitioning to the core insight.'
  }
];

if (typeof module !== 'undefined' && module.exports) {
  module.exports = { PREBUILT_PROMPTS };
}
