// frameworks-data.js — Complete library of 18 pre-built document frameworks for WordAI Pro

const FRAMEWORK_DEFINITIONS = {
  // ─── 🎓 Academic & Scientific Research ─────────────────────────────────────
  'research-proposal': {
    id: 'research-proposal',
    title: '🎓 Academic Research Proposal',
    category: 'academic',
    badge: 'Academic',
    desc: 'Full PhD / Grant proposal blueprint with rationale, PICO questions, methodology matrix, ethics, and 6-month milestone table.',
    sections: ['Title & Abstract', 'Background & Rationale', 'Problem Statement', 'Research Questions & Hypotheses', 'Theoretical Framework', 'Proposed Methodology', 'Ethical Considerations', 'Work Plan & Milestones', 'References'],
    skeleton: `# Title of Research Proposal: [Insert Compelling Project Title]

## Abstract
[Brief 200-word executive overview of the project aims, theoretical contribution, proposed methodology, and societal/scientific significance]

## 1. Background & Rationale
[Provide the broader empirical and theoretical context. Cite foundational literature establishing why this problem matters now]

## 2. Problem Statement & Research Gap
[Precisely articulate the unanswered scientific tension or unexplored empirical phenomenon this study investigates]

## 3. Research Questions & Hypotheses
* **RQ1 (Primary):** [Primary investigative question]
* **RQ2 (Secondary):** [Secondary exploratory question]
* **H1 (Testable Hypothesis):** [Directional hypothesis regarding relationship between independent and dependent variables]

## 4. Theoretical & Conceptual Framework
[Detail the governing theory, conceptual constructs, and causal mechanism diagram/model]

## 5. Proposed Research Methodology
* **Epistemological Stance:** [Positivism / Post-positivism / Interpretivism / Pragmatism]
* **Research Design:** [Mixed-methods sequential explanatory / Longitudinal cohort / Experimental randomized controlled trial]
* **Target Population & Sampling Strategy:** [Sample size n, power analysis calculation, inclusion/exclusion criteria]
* **Instrumentation & Data Collection:** [Standardized scales, structured survey, semi-structured interview protocol]
* **Data Analysis Protocol:** [Inferential statistics (ANCOVA, SEM) or thematic qualitative analysis using NVivo]

## 6. Ethical Safeguards & Data Governance
[IRB / Ethics committee protocol, informed consent, participant anonymization, secure cloud data storage]

## 7. Work Plan & Milestone Schedule
| Milestone / Deliverable | Key Objective | Target Completion | Responsible Lead |
|---|---|---|---|
| Phase 1: Literature Synthesis & Ethics | Finalize literature review & IRB clearance | Month 1-2 | Principal Investigator |
| Phase 2: Instrument Validation & Pilot | Pilot test instruments (n=30) & refine | Month 3 | Research Associate |
| Phase 3: Primary Field Data Collection | Administer instruments across target cohort | Month 4-5 | Lead Researcher |
| Phase 4: Statistical Synthesis & Defense | Analyze dataset, write report & submit | Month 6 | Research Team |

## References
[List foundational peer-reviewed literature in APA 7th format]`,
    aiPrompt: 'You are an elite academic thesis and grant advisor. Write a PhD-grade, highly rigorous Research Proposal following the standard structure: Abstract, Background, Problem Statement, 3 Research Questions, Theoretical Framework, Methodology (design, sampling, measures, data analysis), Ethics, and Timeline.'
  },

  'thesis-5-chapters': {
    id: 'thesis-5-chapters',
    title: '📖 Thesis / Dissertation (Full 5 Chapters)',
    category: 'academic',
    badge: 'Academic',
    desc: 'Comprehensive 5-chapter dissertation architecture: Introduction, Literature Review, Methodology, Results, and Discussion.',
    sections: ['Chapter 1: Introduction', 'Chapter 2: Literature Review', 'Chapter 3: Methodology', 'Chapter 4: Findings & Results', 'Chapter 5: Discussion & Conclusion', 'References & Appendices'],
    skeleton: `# [Title of Dissertation / Master's Thesis]
### By: [Candidate Name] | Department of [Department Name]

---

# Chapter 1: Introduction & Research Problem
## 1.1 Study Context & Justification
[Broad background framing the empirical inquiry]
## 1.2 Statement of the Research Problem
[The core theoretical or practical tension]
## 1.3 Purpose of the Study & Specific Aims
[Explicit research objectives]
## 1.4 Research Questions
* **RQ1:** [Formulate primary question]
* **RQ2:** [Formulate secondary question]
## 1.5 Significance of the Inquiry
[Theoretical, methodological, and practical contributions]

---

# Chapter 2: Review of the Literature
## 2.1 Theoretical Foundations
[Major schools of thought and seminal models]
## 2.2 Thematic Review of Prior Empirical Findings
* **Theme A:** [Synthesize empirical consensus and contradictions]
* **Theme B:** [Emerging developments and contemporary shifts]
## 2.3 Identification of the Research Gap
[Explicitly pinpoint what current literature has overlooked]
## 2.4 Conceptual Model & Hypotheses Development
[Diagram linking constructs to empirical outcomes]

---

# Chapter 3: Research Methodology
## 3.1 Research Philosophy & Paradigm
[Epistemological justification]
## 3.2 Research Design & Strategy
[Justify qualitative, quantitative, or mixed-methods design]
## 3.3 Sampling Frame & Participant Recruitment
[Sample size determination, statistical power, recruitment ethics]
## 3.4 Operationalization of Variables & Instrumentation
[Measurement scales, construct reliability (Cronbach's alpha), validity]
## 3.5 Data Collection & Quality Assurance Procedures
[Field protocol and error reduction measures]
## 3.6 Analytic Techniques
[Step-by-step statistical procedures or qualitative coding frameworks]

---

# Chapter 4: Empirical Results & Findings
## 4.1 Sample Demographics & Descriptive Statistics
| Variable | Mean | Std Dev | Min | Max |
|---|---|---|---|---|
| Sample Cohort A | 0.00 | 0.00 | 0.00 | 0.00 |
| Sample Cohort B | 0.00 | 0.00 | 0.00 | 0.00 |
## 4.2 Hypothesis Testing & Quantitative Syntheses
[Present test statistics (t, F, chi-square, beta), effect sizes, and p-values]
## 4.3 Ancillary & Post-Hoc Analyses
[Robustness checks and sensitivity analyses]

---

# Chapter 5: Discussion, Implications & Conclusions
## 5.1 Synthesis of Findings Relative to Hypotheses
[Reconcile results with Chapter 2 literature]
## 5.2 Theoretical Implications
[How this modifies existing models and paradigms]
## 5.3 Practical & Managerial Recommendations
[Actionable guidelines for industry practitioners or policy makers]
## 5.4 Methodological Limitations & Delimitations
[Acknowledge boundary conditions and constraints]
## 5.5 Directions for Future Scholarly Inquiry
[Propose 3 specific avenues for future research]

---

# References & Appendices
[APA 7th Reference List, Survey Instruments, IRB Approval Documentation]`,
    aiPrompt: 'You are an academic dissertation committee chair. Write an in-depth, rigorous master-level thesis draft following the full 5-chapter architecture: Introduction, Literature Review, Methodology, Results, and Discussion.'
  },

  'prisma-lit-review': {
    id: 'prisma-lit-review',
    title: '📚 Systematic Literature Review (PRISMA 2020)',
    category: 'academic',
    badge: 'Academic',
    desc: 'Gold-standard systematic review protocol following PRISMA guidelines with PICO matrix, search strings, and screening flowchart.',
    sections: ['Rationale & PICO Objectives', 'Eligibility Criteria', 'Search Strategy & Databases', 'Study Selection Protocol', 'Data Extraction Matrix', 'Risk of Bias Assessment', 'Synthesis of Results'],
    skeleton: `# Systematic Literature Review Protocol: [Review Title]
### Adhering to the PRISMA 2020 Statement

## 1. Rationale & PICO Review Objectives
[Contextualize the scientific consensus and state explicit PICO objectives]

| PICO Component | Parameter Description |
|---|---|
| **Population (P)** | [Target demographic, patient cohort, or system profile] |
| **Intervention / Exposure (I)** | [Intervention, algorithm, policy, or treatment applied] |
| **Comparator / Control (C)** | [Alternative standard of care, placebo, or baseline] |
| **Outcomes (O)** | [Primary and secondary quantifiable endpoints] |

## 2. Eligibility Criteria (Inclusion vs. Exclusion)
* **Inclusion Criteria:**
  1. Peer-reviewed journal publications in English (2015–Present)
  2. Empirical studies reporting quantitative or qualitative outcomes matching PICO
  3. Validated measurement methodologies
* **Exclusion Criteria:**
  1. Non-peer reviewed grey literature, editorials, conference abstracts under 2 pages
  2. Studies lacking verifiable methodological data or control conditions

## 3. Information Sources & Search Strategy Query String
* **Databases Polled:** PubMed/MEDLINE, Scopus, Web of Science, IEEE Xplore, JSTOR.
* **Exact Search Syntax String:**
\`\`\`text
("Primary Term A" OR "Synonym A1") AND ("Intervention B" OR "Approach B1") AND ("Outcome C" OR "Measure C1")
\`\`\`

## 4. Study Selection & Screening Protocol
[Describe the 2-reviewer independent screening procedure for title/abstract screening followed by full-text review. Detail Cohen's Kappa score for inter-rater agreement]

## 5. Data Extraction Form & Synthesis Protocol
| Study (Author, Year) | Sample Size (n) | Methodology / Design | Key Finding / Effect Size | Quality Rating |
|---|---|---|---|---|
| Author A et al. (2023) | n=150 | Randomized Trial | Significant improvement (d=0.65) | High |
| Author B et al. (2022) | n=85 | Quasi-Experimental | Moderate correlation (r=0.42) | Medium |

## 6. Risk of Bias & Methodological Quality Assessment
[Detail application of Cochrane RoB 2, Newcastle-Ottawa Scale, or ROBINS-I tool across bias domains]

## 7. Narrative & Meta-Analytic Synthesis of Findings
[Synthesize findings across thematic clusters, explore heterogeneity, and declare conclusive consensus]`,
    aiPrompt: 'You are a meta-analysis and systematic review expert. Draft a PRISMA-compliant Systematic Literature Review protocol complete with PICO formulation, boolean search strings, eligibility criteria, and data synthesis tables.'
  },

  'imrad-article': {
    id: 'imrad-article',
    title: '🔬 IMRaD Scientific Journal Article',
    category: 'academic',
    badge: 'Academic',
    desc: 'Standard international scientific journal format: Introduction, Methods, Results, and Discussion with structured abstract.',
    sections: ['Title & Structured Abstract', 'Introduction', 'Materials & Methods', 'Results', 'Discussion', 'Conclusions & References'],
    skeleton: `# [Manuscript Title: Informative, Concise, Highlighting Primary Finding]

## Structured Abstract
* **Background:** [1-2 sentences framing the research problem]
* **Methods:** [2 sentences detailing experimental design and sample]
* **Results:** [2-3 sentences providing exact quantitative findings with statistical metrics]
* **Conclusions:** [1-2 sentences stating the definitive scientific contribution]
* **Keywords:** [5-6 MeSH / Index Keywords]

## 1. Introduction
[Hook the reader with scientific context. Transition into known literature, pinpoint the unresolved conflict, and state the formal research hypothesis]

## 2. Materials & Methods
### 2.1 Experimental Cohort & Sampling
[Detailed participant / material criteria and ethical approval numbers]
### 2.2 Experimental Protocol & Apparatus
[Step-by-step description allowing exact reproducibility by another researcher]
### 2.3 Statistical Analysis Framework
[Power analysis, normality tests, inferential models, significance threshold alpha=0.05]

## 3. Results
### 3.1 Primary Outcome Findings
[Report data objectively without interpretive bias]
| Experimental Group | Baseline Mean (SD) | Post-Intervention (SD) | p-value | Effect Size (Cohen's d) |
|---|---|---|---|---|
| Control Cohort (n=50) | 12.4 (2.1) | 12.6 (2.3) | p > 0.05 | 0.09 |
| Treatment Cohort (n=50)| 12.2 (2.0) | 18.7 (1.8) | p < 0.001 | 1.42 |

### 3.2 Secondary Measures & Sensitivity Analysis
[Report secondary metrics, subgroup effects, and covariate controls]

## 4. Discussion
[Interpret findings immediately relative to the initial hypothesis. Compare with prior published findings. Explain physiological or mechanistic underpinnings]
* **Study Strengths:** [Rigorous controls, validated instruments, representative sample]
* **Study Limitations:** [Sample delimitations, confounding variables, observational limits]

## 5. Conclusions & Scientific Implications
[Final punchy takeaway and translational significance for future research]

## References
[APA / Nature / IEEE formatted citation index]`,
    aiPrompt: 'You are a senior scientific editor for a top peer-reviewed journal. Write a rigorous IMRaD scientific manuscript with structured abstract, empirical methodology, statistical results tables, and high-level discussion.'
  },

  'lab-report': {
    id: 'lab-report',
    title: '🧪 Scientific Lab Investigation Report',
    category: 'academic',
    badge: 'Academic',
    desc: 'Laboratory experiment report with physical equations, precision tolerances, data tables, error analysis, and physical discussion.',
    sections: ['Objective & Theoretical Principles', 'Apparatus & Equipment Tolerances', 'Experimental Procedure', 'Data Observations Table', 'Calculations & Percent Error', 'Discussion of Error Sources', 'Conclusions'],
    skeleton: `# Laboratory Investigation Report: [Title of Experiment]
### Course / Module: [Course Code] | Station / Group: [Lab Group #]

## Abstract
[Quantitative summary of hypotheses tested, experimental protocol, measured constants, percent error, and main conclusion]

## 1. Objectives & Theoretical Principles
* **Primary Objective:** [State exact experimental purpose]
* **Governing Mathematical Equations:**
  $$F = m \\cdot a \\quad \\text{or} \\quad V = I \\cdot R$$
* **Theoretical Predictions:** [Anticipated linear, exponential, or inverse relationship]

## 2. Apparatus, Materials & Measurement Precision
| Equipment / Instrument | Model / Serial | Measurement Range | Instrument Precision / Tolerance |
|---|---|---|---|
| Digital Micrometer | Mitutoyo 293 | 0 - 25 mm | ± 0.001 mm |
| Digital Balance | Ohaus Adventurer | 0 - 220 g | ± 0.0001 g |
| Digital Multimeter | Fluke 87V | 0 - 1000 V | ± 0.05% + 1 digit |

## 3. Experimental Procedure
1. [Initial calibration and zero-offset recording]
2. [Sequential variation of independent variable]
3. [Repetition protocol to obtain 3 trials per measurement step]
4. [Safety precautions and environmental control parameters]

## 4. Experimental Data & Observations
| Trial # | Independent Variable (X) | Dependent Trial 1 | Dependent Trial 2 | Dependent Trial 3 | Mean Value |
|---|---|---|---|---|---|
| 1 | 10.0 units | 20.2 | 20.4 | 20.1 | 20.23 units |
| 2 | 20.0 units | 40.8 | 40.5 | 40.9 | 40.73 units |
| 3 | 30.0 units | 61.1 | 60.8 | 61.3 | 61.07 units |

## 5. Quantitative Calculations & Error Propagation
* **Sample Calculation:** [Demonstrate one complete calculation from raw data to final value]
* **Theoretical Accepted Value:** [Literature value]
* **Measured Experimental Value:** [Calculated mean value]
* **Percent Error Calculation:**
  $$\\text{Percent Error} = \\frac{|\\text{Experimental} - \\text{Theoretical}|}{\\text{Theoretical}} \\times 100\\% = 2.45\\%$$

## 6. Discussion of Systematic & Random Error
* **Systematic Error Sources:** [Zero drift, calibration offset, heat dissipation losses]
* **Random Error Sources:** [Parallax observation error, ambient temperature fluctuations]
* **Methodological Recommendations:** [How to eliminate these errors in future trials]

## 7. Conclusions
[Final synthesis confirming or rejecting the theoretical relationship within experimental uncertainty]`,
    aiPrompt: 'You are a laboratory physicist and scientific researcher. Draft a complete, mathematically precise Laboratory Report including theory, apparatus table, raw data trial tables, error propagation formulas, and error source discussion.'
  },

  'academic-case-study': {
    id: 'academic-case-study',
    title: '🏛️ Academic Case Study & Qualitative Analysis',
    category: 'academic',
    badge: 'Academic',
    desc: 'Deep qualitative case analysis featuring institutional context, critical incidents, data triangulation, and pedagogical discussion.',
    sections: ['Executive Overview & Case Profile', 'Context & Chronological Narrative', 'The Critical Dilemma', 'Qualitative Triangulation', 'Thematic Analysis & Coding', 'Pedagogical / Policy Recommendations'],
    skeleton: `# Academic Case Study: [Organization / Subject Name]
### Subtitle: [The Central Strategic, Clinical, or Societal Dilemma]

## 1. Executive Overview & Case Profile
[Summary of the institutional setting, central protagonist, timeline, and core decisions facing leadership]

## 2. Institutional Context & Chronological Background
* **Organization Profile:** [History, mission, organizational structure, market/clinical niche]
* **The Operating Environment:** [Regulatory pressures, socio-economic forces, stakeholder landscape]
* **Key Historical Milestones Leading to the Crisis:** [Timeline of precursor events]

## 3. The Critical Incident & Decision Dilemma
[Narrate the pivotal moment where existing protocols failed or conflicting interests collided. Detail the high-stakes trade-offs]

## 4. Qualitative Data Triangulation
| Evidence Stream | Source Description | Key Corroborating Insight | Potential Bias / Delimitation |
|---|---|---|---|
| Stakeholder Interviews | n=12 Executive and frontline staff | Unanimous recognition of communication breakdown | Retrospective recall bias |
| Internal Documentation | Internal email correspondence & audit logs | Decision made despite engineer warnings | Incomplete documentation |
| Third-Party Observation | Independent regulatory inspection | Systemic non-compliance with safety codes | Point-in-time snapshot |

## 5. Thematic Analysis & Theoretical Cross-Examination
### Theme 1: [Identify Core Theoretical Failure or Mechanism]
[Analyze the phenomenon through established scholarly lenses]
### Theme 2: [Incentive Misalignment or Cultural Pathology]
[Examine underlying organizational drivers]

## 6. Pedagogical Discussion Questions & Policy Recommendations
1. [Discussion Question 1 challenging assumptions]
2. [Discussion Question 2 forcing students to prioritize conflicting values]
* **Strategic Policy Interventions:** [Itemize 3 actionable, evidence-based recommendations to prevent recurrence]`,
    aiPrompt: 'You are a professor of organizational sociology and business policy. Write an academic teaching case study complete with institutional profile, narrative dilemma, qualitative data matrix, thematic coding analysis, and pedagogical discussion questions.'
  },

  // ─── 💼 Business, Strategy & Management ────────────────────────────────────
  'business-plan': {
    id: 'business-plan',
    title: '💼 Comprehensive Executive Business Plan',
    category: 'business',
    badge: 'Business',
    desc: 'Bank-ready & venture-ready business plan: TAM/SAM/SOM market sizing, competitor matrix, unit economics, and 3-year P&L projections.',
    sections: ['Executive Summary & Mission', 'Problem & Value Proposition', 'Market Analysis & TAM/SAM/SOM', 'Competitive Advantage Matrix', 'Product & Technical Roadmap', 'Marketing & Acquisition Strategy', 'Financial Model & 3-Year Projections', 'The Ask & Use of Funds'],
    skeleton: `# Comprehensive Business Plan: [Company / Venture Name]
### Confidentially Prepared for Investors & Financial Partners

## 1. Executive Summary
* **Mission Statement:** [One compelling sentence stating company purpose]
* **The Problem:** [Clear friction point costing customers time, money, or efficiency]
* **The Solution:** [Proprietary product or service delivering order-of-magnitude improvement]
* **Company Status & Traction:** [Founding date, current revenue run-rate, active users, key IP]
* **The Ask:** [Seeking $X Million in funding for Y months of runway]

## 2. Market Opportunity & Sizing
* **Total Addressable Market (TAM):** [$X Billion worldwide market]
* **Serviceable Addressable Market (SAM):** [$Y Million targeted segment]
* **Serviceable Obtainable Market (SOM):** [$Z Million realistically achievable within 3 years]
* **Market Megatrends & Growth Drivers:** [3 macro forces propelling adoption]

## 3. Competitive Landscape & Advantage Matrix
| Competitor | Market Share | Pricing Model | Key Vulnerability | Our Competitive Moat |
|---|---|---|---|---|
| Incumbent Giant | 55% | High Enterprise Contract | Legacy slow software | 10x faster self-serve UX |
| Low-End Challenger | 15% | Free / Freemium | No enterprise security | SOC-2 Type II + SLA guarantee |
| **[Our Company]** | Growing | Subscription / Tiered | Emerging brand | Proprietary AI workflow engine |

## 4. Product, Service Line & Technology Architecture
[Detail the core features, technical architecture, proprietary IP/patents, and future R&D roadmap]

## 5. Go-To-Market, Sales & Distribution Strategy
* **Customer Acquisition Channels:** [Inbound SEO, enterprise sales reps, strategic partnerships]
* **Target Unit Economics:**
  * Customer Acquisition Cost (CAC): [$X]
  * Lifetime Value (LTV): [$Y]
  * LTV : CAC Ratio: [Target 3:1 or 4:1]
  * Payback Period: [Target < 12 Months]

## 6. Financial Model & 3-Year Projections
| Financial Metric | Year 1 | Year 2 | Year 3 |
|---|---|---|---|
| **Total Revenue** | $500,000 | $2,200,000 | $7,500,000 |
| **Gross Margin %** | 72% | 78% | 82% |
| **Operating Expenses (OPEX)** | $650,000 | $1,800,000 | $4,500,000 |
| **EBITDA** | ($150,000) | $400,000 | $3,000,000 |
| **Full-Time Headcount** | 6 FTE | 18 FTE | 45 FTE |

## 7. The Ask, Funding Allocation & Key Milestones
* **Capital Requirement:** Seeking [$X,000,000]
* **Use of Proceeds:** 50% Engineering & Product, 35% Sales & Marketing, 15% Operations & Legal.
* **18-Month Success Milestones:** [3 verifiable targets achieved with this capital]`,
    aiPrompt: 'You are a veteran venture capital investor and business strategist. Draft a comprehensive, bankable executive business plan complete with TAM/SAM/SOM market sizing, competitive matrix, unit economics, and 3-year financial model.'
  },

  'hbs-case-analysis': {
    id: 'hbs-case-analysis',
    title: '📈 Case Study Analysis (Harvard Business Format)',
    category: 'business',
    badge: 'Business',
    desc: 'Harvard Business School executive analysis: Organizational diagnosis, root cause pathology, strategic alternatives, and 90-day action plan.',
    sections: ['Executive Summary', 'Macro & Industry Diagnostic', 'Root Cause Analysis', 'Decision Criteria & Constraints', 'Strategic Alternatives Evaluation', 'Recommended Strategy', '90-Day Tactical Implementation'],
    skeleton: `# Strategic Case Analysis: [Company Name]
### Harvard Business School Analytic Format

## 1. Executive Summary
[Synthesize the critical strategic decision facing the CEO/Board, your core recommendation, and the expected ROI within 180 words]

## 2. Organizational & Industry Diagnostic
* **Internal Capabilities (VRIO Framework):** [Valuable, Rare, Inimitable, Organized resources]
* **Industry Structure (Porter's Five Forces):**
  * Threat of New Entrants: [High/Med/Low with justification]
  * Bargaining Power of Buyers: [High/Med/Low with justification]
  * Threat of Substitutes: [High/Med/Low with justification]
  * Intensity of Rivalry: [High/Med/Low with justification]

## 3. Root Cause Analysis (Symptoms vs. Underlying Pathology)
* **Surface Symptoms:** [Declining gross margins, executive turnover, customer complaints]
* **Underlying Root Pathology:** [Structural mispricing, misaligned sales incentives, technical debt]

## 4. Strategic Decision Criteria & Boundaries
1. **Strategic Fit:** [Aligns with core competencies]
2. **Financial Hurdle:** [Internal Rate of Return > 20%, Payback < 24 months]
3. **Execution Feasibility:** [Can be deployed with current engineering bandwidth]
4. **Brand / Regulatory Risk:** [Preserves enterprise customer trust]

## 5. Evaluation of Strategic Alternatives
| Strategic Option | Pros | Cons / Risks | Hurdle Score (1-10) |
|---|---|---|---|
| **Option A: Retrench & Protect Core** | Low execution risk, immediate cash flow | Surrenders future growth market | 6.5 / 10 |
| **Option B: Aggressive M&A Acquisition** | Immediate market share and IP access | High integration failure risk | 7.0 / 10 |
| **Option C: Digital Platform Pivot** | High long-term margins, recurring revenue | Requires cultural re-skilling | 9.0 / 10 (Winner) |

## 6. Recommended Strategic Path & Value Creation Rationale
[Articulate why Option C is superior. Present quantitative sensitivity analysis proving robustness]

## 7. 90-Day Tactical Implementation & Governance Roadmap
| Phase | Weeks | Critical Action Milestones | Risk Mitigation Measure |
|---|---|---|---|
| **Phase 1: Mobilize** | Weeks 1-4 | Establish Transformation Office & align executives | Clear OKR incentives |
| **Phase 2: Pilot** | Weeks 5-8 | Roll out alpha program to top 10 enterprise clients | Dedicated white-glove support |
| **Phase 3: Scale** | Weeks 9-12 | Full commercial launch & decommissioning legacy tool | Automated migration tools |`,
    aiPrompt: 'You are a McKinsey senior partner and Harvard Business School professor. Analyze this business case using the classic HBS framework: Executive Summary, Porter 5 Forces diagnostic, root-cause analysis, 3 strategic options with trade-off matrix, and 90-day implementation roadmap.'
  },

  'project-charter-sow': {
    id: 'project-charter-sow',
    title: '📋 Project Charter & Scope of Work (SOW)',
    category: 'business',
    badge: 'Business',
    desc: 'Formal enterprise Project Charter and Scope of Work (SOW) with RACI matrix, in/out scope boundaries, milestones, and risk register.',
    sections: ['Project Vision & Business Case', 'Scope Boundaries (In vs Out)', 'Key Deliverables & Acceptance Criteria', 'RACI Responsibility Matrix', 'Milestone Schedule & Budget', 'Assumptions & Risk Register'],
    skeleton: `# Project Charter & Statement of Work: [Project Title]
### Project Code: [PRJ-YYYY-###] | Sponsor: [Executive Sponsor Name]

## 1. Project Purpose & Strategic Justification
* **Project Vision:** [Clear description of intended business transformation]
* **Business Need & ROI:** [Expected operational cost reduction, revenue unlock, or compliance mandate]
* **Success Criteria:** [3 measurable quantitative metrics defining project success]

## 2. Scope Boundaries (Explicit In-Scope vs. Out-of-Scope)
| In-Scope (Committed Deliverables) | Out-of-Scope (Explicit Exclusions) |
|---|---|
| Cloud infrastructure migration to AWS | On-premise hardware maintenance |
| Data migration of active customer records (2020–Present) | Archival records pre-dating 2020 |
| Single Sign-On (SSO) integration via Okta | Custom biometric multi-factor authentication |
| End-user training workshops (3 sessions) | Ongoing Level-1 technical helpdesk support |

## 3. Key Deliverables & Acceptance Criteria
* **Deliverable 1:** [Specification, verification protocol, and sign-off authority]
* **Deliverable 2:** [Specification, verification protocol, and sign-off authority]

## 4. RACI Governance & Stakeholder Matrix
*(R = Responsible, A = Accountable, C = Consulted, I = Informed)*

| Project Role / Stakeholder | System Architecture | Data Migration | User Acceptance Testing | Go-Live Decision |
|---|---|---|---|---|
| Project Manager | A | A | A | R |
| Lead Technical Architect | R | C | C | C |
| Business Unit Lead | C | C | R | A |
| InfoSec / Compliance Officer | C | I | C | C |

## 5. Milestone Schedule & Budget Allocation
| Milestone Phase | Planned Start | Target Sign-Off | Budget Allocation ($) |
|---|---|---|---|
| Discovery & Architecture Blueprint | Month 1, Week 1 | Month 1, Week 4 | $25,000 |
| Core Development & Integration | Month 2, Week 1 | Month 3, Week 4 | $80,000 |
| User Acceptance Testing (UAT) | Month 4, Week 1 | Month 4, Week 3 | $20,000 |
| Production Go-Live & Handover | Month 4, Week 4 | Month 5, Week 1 | $15,000 |

## 6. Assumptions, Dependencies & Risk Log
| Risk Event Description | Probability (1-5) | Impact (1-5) | Mitigation / Contingency Plan |
|---|---|---|---|
| Delays in API access from legacy vendor | 4 | 4 | Pre-negotiate SLA penalties & build mock data servers |
| Key technical personnel turnover | 2 | 4 | Comprehensive architectural documentation & cross-training |`,
    aiPrompt: 'You are a PMP-certified senior enterprise project director. Draft a formal, ironclad Project Charter and Statement of Work (SOW) complete with scope boundaries table, deliverables acceptance criteria, RACI matrix, and risk register.'
  },

  'standard-operating-procedure': {
    id: 'standard-operating-procedure',
    title: '⚙️ Standard Operating Procedure (SOP)',
    category: 'business',
    badge: 'Business',
    desc: 'Audit-ready SOP protocol: Regulatory compliance standards, role duties, step-by-step procedure, failure modes, and revision logs.',
    sections: ['Purpose & Scope', 'Regulatory Governance & Standards', 'Roles & Responsibilities', 'Step-by-Step Procedure', 'Contingency & Failure Recovery', 'Revision History & Sign-Off'],
    skeleton: `# Standard Operating Procedure: [Procedure Title]
### Document ID: [SOP-DEPT-###] | Version: 1.0 | Effective Date: [Date]

## 1. Purpose & Operational Objectives
[State the exact operational purpose of this procedure, safety imperatives, and quality standards achieved by adherence]

## 2. Scope & Applicability
* **Applies to:** [Specific departments, facilities, job roles, or systems]
* **Exclusions:** [Scenarios where alternate emergency protocols supersede this document]

## 3. Regulatory Governance & Reference Standards
* ISO 9001:2015 Quality Management Systems (Clause X)
* OSHA / Regulatory Compliance Guidelines (Standard Y)
* Internal Corporate Governance Policy (Ref Z)

## 4. Roles, Qualifications & Responsibilities
| Role | Required Qualifications / Training | Operational Responsibilities |
|---|---|---|
| Operator / Technician | Certified Level-2 Operator | Execution of steps 5.1 through 5.5 |
| Shift Supervisor | Lead Quality Certification | Verification, log sign-off, exception escalation |
| Quality Auditor | Internal Auditor Certified | Bi-weekly random sampling and calibration checks |

## 5. Step-by-Step Execution Protocol
### 5.1 Pre-Execution Preparation & Safety Verification
1. Verify personal protective equipment (PPE) is inspected and worn.
2. Confirm instrument calibration date is active and within tolerance.
3. Review job safety hazard analysis (JHA) log.

### 5.2 Primary Procedural Execution
| Step # | Action Description | Target Specification | Corrective Action if Out-of-Spec |
|---|---|---|---|
| 1 | Initialize baseline system check | Green status indicator | Cycle power; notify supervisor if red persists |
| 2 | Feed raw material input | Batch temp between 20°C–24°C | Adjust thermal jacket before processing |
| 3 | Execute primary operational cycle | Cycle time: 180s ± 5s | Halt cycle if pressure exceeds 45 PSI |

### 5.3 Post-Execution Cleaning, Logging & Handover
1. Purge system chambers and log waste volume in compliance ledger.
2. Perform electronic sign-off in Enterprise Resource Planning (ERP) database.

## 6. Contingency, Exception & Failure Recovery Protocols
[Detail immediate emergency shutdown protocols, quarantine procedures for non-conforming product, and incident reporting deadlines]

## 7. Revision History, Review Schedule & Approvals
| Version | Date | Section Changed | Reason for Modification | Approved By |
|---|---|---|---|---|
| 1.0 | [Date] | Initial Release | Baseline operational standard | Quality Director |`,
    aiPrompt: 'You are an ISO 9001 quality director and compliance engineer. Write an audit-ready Standard Operating Procedure (SOP) with regulatory governance, safety controls, step-by-step procedural matrix, and exception handling.'
  },

  'quarterly-business-review': {
    id: 'quarterly-business-review',
    title: '📊 Quarterly Business Review (QBR)',
    category: 'business',
    badge: 'Business',
    desc: 'Executive QBR presentation document: OKR performance scorecard, revenue variance analysis, retrospective post-mortem, and next quarter priorities.',
    sections: ['Executive Highlights & Top Wins', 'OKR & Core KPI Scorecard', 'Financial Performance & Revenue Variance', 'Retrospective Lessons & Blockers', 'Next Quarter Strategic Priorities'],
    skeleton: `# Quarterly Business Review: [Quarter / Year, e.g. Q3 2026]
### Department / Business Unit: [Unit Name] | Presented to: [Executive Leadership Team]

## 1. Executive Highlights & Top Strategic Wins
* **Headline Achievement:** [Top strategic accomplishment achieved this quarter]
* **Customer / Market Milestone:** [Key enterprise logo closed or product launched]
* **Operational Efficiency Win:** [Cost reduction or cycle-time compression metric]

## 2. OKR & Core KPI Performance Scorecard
| Objective / KPI | Quarter Target | Actual Result | Status | Variance % |
|---|---|---|---|---|
| Annual Recurring Revenue (ARR) | $4.50M | $4.82M | 🟢 Exceeded | +7.1% |
| Net Revenue Retention (NRR) | 115% | 118% | 🟢 Exceeded | +3.0% |
| Customer Churn Rate | < 1.2% | 1.8% | 🔴 Off-Track | +0.6% Churn |
| Product NPS Score | 60 | 64 | 🟢 On-Track | +4 pts |

## 3. Financial Performance & P&L Variance Commentary
* **Revenue Drivers:** [Identify which product lines, customer cohorts, or regions outperformed]
* **Expense Variance:** [Explain budget overages or headcount timing savings]
* **Gross Margin Evolution:** [Analysis of infrastructure costs and COGS efficiency]

## 4. Retrospective Lessons, Churn Drivers & Strategic Blockers
* **What Went Wrong:** [Unvarnished analysis of missed targets or delayed releases]
* **Root Causes for Customer Churn:** [Product feature gaps, onboarding friction, competitor price cuts]
* **Remediation Implemented:** [Specific fixes deployed to prevent recurrence]

## 5. Next Quarter Strategic Priorities & Resource Allocation
| Priority Rank | Strategic Objective | Expected Business Impact | Key Dependencies |
|---|---|---|---|
| 1 | Enterprise Security & SOC-2 Certification | Unlocks $1.5M blocked enterprise pipeline | InfoSec engineering support |
| 2 | Automated Self-Serve Onboarding Flow | Reduces time-to-value from 14 days to 4 hours | Product & UX Design team |
| 3 | Europe / EMEA Geographic Expansion | Expands addressable TAM by 35% | Legal GDPR compliance sign-off |

## 6. Executive Decision & Resource Requests
[Explicit requests for budget authorization, cross-functional engineering headcount, or executive sponsorship]`,
    aiPrompt: 'You are a Chief of Staff and VP of Corporate Strategy. Draft an executive Quarterly Business Review (QBR) document featuring OKR scorecards, revenue variance analysis, churn retrospective, and prioritized strategic initiatives for next quarter.'
  },

  'executive-whitepaper': {
    id: 'executive-whitepaper',
    title: '📄 Executive White Paper & Industry Report',
    category: 'business',
    badge: 'Business',
    desc: 'Authoritative thought-leadership report: Market megatrends, structural inefficiencies, modern solution framework, and enterprise ROI proof.',
    sections: ['Executive Summary', 'Market Dynamics & Industry Megatrends', 'The Structural Inefficiency', 'The Architectural Solution Framework', 'Enterprise Case Studies & ROI', 'Strategic Roadmap & Checklist'],
    skeleton: `# [White Paper Title: Authoritative, Insightful, Pointing to Future State]
### An Executive Briefing on [Emerging Paradigm / Technology Transformation]

## Executive Summary & Core Thesis
[A compelling 250-word synthesis of why traditional approaches are failing, the emergence of the new architectural paradigm, and the quantified business impact of early adoption]

## 1. Market Dynamics & The Shift in Industry Megatrends
[Analyze macro-economic, regulatory, or technological shifts that have rendered legacy operating models obsolete. Cite reputable industry benchmarks (Gartner, Forrester, McKinsey)]

## 2. The Core Friction: Structural Inefficiencies in Current Paradigms
* **Legacy Bottleneck A:** [Explain why current workflows collapse under modern scale]
* **Hidden Economic Costs:** [Quantify lost employee productivity, security vulnerabilities, or customer churn]
* **The High Cost of Inaction:** [What happens to companies that maintain the status quo over 24 months]

## 3. The Modern Solution Framework: Next-Generation Architecture
[Introduce your proposed methodology, operational model, or technological framework]
| Legacy Operational Paradigm | The Modern Architectural Framework | Strategic Advantage |
|---|---|---|
| Fragmented manual data silos | Real-time unified AI data mesh | Zero latency intelligence |
| Reactive batch incident response | Proactive autonomous self-healing | 99.99% system availability |
| High fixed personnel overhead | Scalable API-driven automation | 60% reduction in OPEX |

## 4. Empirical Case Studies & Quantified Business Impact
### Enterprise Case Study: [Fortune 500 Global Corporation]
* **Initial Challenge:** [Baseline inefficiency costing millions]
* **Implementation Strategy:** [Phased rollout across 60 days]
* **Quantified ROI Results:**
  * 42% decrease in customer onboarding duration
  * $3.2 Million annual cost avoidance
  * 18-month projected net payback achieved in 4.5 months

## 5. Strategic Roadmap: Evaluation & Implementation Checklist
1. **Assessment Phase (Days 1–30):** [Audit current technical debt and integration dependencies]
2. **Pilot Deployment (Days 31–60):** [Run parallel sandbox testing on high-value workflow]
3. **Enterprise Rollout (Days 61–90):** [Decommission legacy tool and scale enterprise-wide]

## Conclusion & Next Steps
[Call to action for enterprise leaders and author / advisory contact information]`,
    aiPrompt: 'You are an enterprise CTO and thought-leadership consultant. Draft an authoritative, C-suite executive white paper detailing market megatrends, legacy structural bottlenecks, a modern solution framework, and empirical ROI metrics.'
  },

  // ─── ⚖️ Legal, Policy & Governance ─────────────────────────────────────────
  'mutual-nda': {
    id: 'mutual-nda',
    title: '🔒 Mutual Non-Disclosure Agreement (NDA)',
    category: 'legal',
    badge: 'Legal',
    desc: 'Standard commercial bilateral NDA: Definition of confidential information, duty of care, carve-outs, return of materials, and governing law.',
    sections: ['Preamble & Purpose', 'Definition of Confidential Information', 'Obligations & Standard of Care', 'Carve-Outs & Exclusions', 'Term & Return of Materials', 'Remedies & Governing Jurisdiction'],
    skeleton: `# MUTUAL NON-DISCLOSURE AND CONFIDENTIALITY AGREEMENT

This Mutual Non-Disclosure Agreement ("Agreement") is entered into as of **[Effective Date]**, by and between:

* **Party A:** [Company Name A], a [State/Country] Corporation, with its principal place of business at [Address A] ("Party A"); and
* **Party B:** [Company Name B], a [State/Country] Corporation, with its principal place of business at [Address B] ("Party B").

*(Each a "Party" and collectively the "Parties").*

## 1. Purpose
The Parties desire to explore a mutually beneficial business opportunity relating to **[Describe Purpose of Discussions, e.g. a potential strategic partnership or software integration]** (the "Purpose"). In connection therewith, each Party may disclose proprietary confidential information to the other.

## 2. Definition of Confidential Information
"Confidential Information" means all non-public, confidential, or proprietary information disclosed by one Party ("Disclosing Party") to the other Party ("Receiving Party"), whether orally, visually, or in tangible/electronic form, including but not limited to:
* Technical data, source code, software architecture, algorithms, trade secrets, inventions, and know-how;
* Financial statements, customer lists, pricing strategies, business plans, and commercial terms;
* Any information that should reasonably be understood to be confidential given its nature and circumstances of disclosure.

## 3. Obligations & Standard of Care
The Receiving Party agrees:
1. To hold the Disclosing Party's Confidential Information in strict confidence, applying at least the same degree of care it uses to protect its own confidential information of like kind, but no less than a reasonable degree of care;
2. Not to disclose such Confidential Information to any third party, except to its officers, directors, employees, and legal/financial advisors who have a strict need to know for the Purpose and are bound by written non-disclosure obligations at least as restrictive as this Agreement;
3. Not to use the Disclosing Party's Confidential Information for any purpose other than the authorized Purpose without prior written consent.

## 4. Exclusions from Confidential Treatment
Confidential Information does not include information that:
* Is or becomes publicly known through no breach of this Agreement by the Receiving Party;
* Was already lawfully in the possession of the Receiving Party without restriction prior to disclosure;
* Is independently developed by the Receiving Party without reference to or reliance upon the Disclosing Party's Confidential Information;
* Is rightfully received from a third party without an obligation of confidentiality.

## 5. Term, Termination & Survival
This Agreement shall govern disclosures made for a period of **two (2) years** from the Effective Date. The confidentiality obligations under this Agreement shall survive for a period of **three (3) years** following expiration or termination, provided that trade secrets shall remain confidential in perpetuity.

## 6. Return or Destruction of Materials
Promptly upon written request of the Disclosing Party, the Receiving Party shall return or destroy all physical and electronic documents containing Confidential Information, and certify compliance in writing.

## 7. Equitable Remedies & Governing Law
The Parties acknowledge that any breach of this Agreement may cause irreparable harm for which monetary damages alone would be inadequate, and agree that the Disclosing Party shall be entitled to seek injunctive relief. This Agreement shall be governed by the laws of **[State/Country, e.g. the State of Delaware]**, without regard to conflict of law principles.

**IN WITNESS WHEREOF**, the Parties have executed this Agreement by their duly authorized representatives.

| Party A: [Company Name A] | Party B: [Company Name B] |
|---|---|
| By: __________________________ | By: __________________________ |
| Name: [Authorized Signatory] | Name: [Authorized Signatory] |
| Title: [Title] | Title: [Title] |
| Date: [Date] | Date: [Date] |`,
    aiPrompt: 'You are corporate legal counsel. Draft a comprehensive, bilateral Mutual Non-Disclosure Agreement (NDA) with standard commercial protections, precise definitions, carve-outs, survival terms, and signature blocks.'
  },

  'consulting-agreement': {
    id: 'consulting-agreement',
    title: '📑 Master Services & Consulting Agreement',
    category: 'legal',
    badge: 'Legal',
    desc: 'Independent contractor agreement: SOW deliverables, compensation, intellectual property work-for-hire assignment, liability, and termination.',
    sections: ['Services & SOW Deliverables', 'Compensation & Invoicing', 'Intellectual Property Assignment', 'Independent Contractor Status', 'Warranties & Liability', 'Term & Termination'],
    skeleton: `# MASTER CONSULTING SERVICES AGREEMENT (MSA)

This Consulting Agreement ("Agreement") is made effective as of **[Effective Date]**, by and between:
* **Client:** [Client Company Name] ("Client"), located at [Client Address]; and
* **Consultant:** [Consultant Name / Firm] ("Consultant"), located at [Consultant Address].

## 1. Services & Statements of Work
Consultant agrees to provide professional consulting services as specified in one or more Statements of Work ("SOW") executed by both Parties. Each SOW shall detail project scope, deliverables, acceptance criteria, and specific milestone timelines.

## 2. Compensation & Invoicing Schedule
* **Fee Structure:** Client shall pay Consultant according to the fee schedule specified in each SOW (e.g. fixed fee of [$X] or hourly rate of [$Y/hr]).
* **Payment Terms:** Consultant shall invoice Client monthly. Invoices are payable within **thirty (30) days** (Net 30) of invoice receipt.
* **Reimbursable Expenses:** Pre-approved out-of-pocket travel and lodging expenses shall be reimbursed at cost upon presentation of valid receipts.

## 3. Intellectual Property Rights & Work Made for Hire
* **Assignment of Work Product:** Consultant agrees that all inventions, designs, source code, deliverables, and works of authorship created for Client under this Agreement shall constitute "works made for hire".
* To the extent any deliverable does not qualify as a work made for hire, Consultant hereby unconditionally assigns to Client all right, title, and interest, including worldwide copyright and patent rights, in and to such Work Product.
* **Pre-Existing IP:** Consultant retains ownership of pre-existing tools and generic methodologies, granting Client a perpetual, royalty-free license to use such pre-existing IP embedded in the deliverables.

## 4. Independent Contractor Status
Consultant is an independent contractor and not an employee, agent, or partner of Client. Consultant is solely responsible for all income taxes, withholdings, social security, and health insurance obligations.

## 5. Warranties & Limitation of Liability
* Consultant warrants that all services will be performed in a professional, workmanlike manner conforming to prevailing commercial standards.
* **Cap on Liability:** Except for breaches of confidentiality or gross negligence, neither party's total aggregate liability under this Agreement shall exceed the total fees paid by Client to Consultant under the applicable SOW in the preceding 12 months.

## 6. Term & Termination
* **Termination for Convenience:** Either party may terminate this Agreement without cause upon **thirty (30) days' written notice**.
* **Termination for Cause:** Either party may terminate immediately upon written notice if the other party materially breaches this Agreement and fails to cure within fifteen (15) days.

| Client: [Client Company Name] | Consultant: [Consultant / Firm Name] |
|---|---|
| By: __________________________ | By: __________________________ |
| Name: [Signatory Name] | Name: [Signatory Name] |
| Date: [Date] | Date: [Date] |`,
    aiPrompt: 'You are corporate legal counsel. Draft a robust Master Services & Independent Consulting Agreement including work-for-hire IP assignments, invoicing schedules, independent contractor protections, and liability limits.'
  },

  'privacy-policy-gdpr': {
    id: 'privacy-policy-gdpr',
    title: '🛡️ Privacy Policy & GDPR/CCPA Notice',
    category: 'legal',
    badge: 'Legal',
    desc: 'Comprehensive data privacy policy compliant with GDPR and CCPA: Data collection categories, legal bases, sub-processors, and user rights.',
    sections: ['Data Controller Identification', 'Categories of Data Collected', 'Legal Grounds for Processing', 'Third-Party Processors & Sharing', 'International Data Transfers', 'User Rights & Retention Policies'],
    skeleton: `# PRIVACY POLICY & DATA PROTECTION NOTICE
### Last Updated: [Date, e.g. September 2026]

## 1. Introduction & Controller Identification
**[Company / Service Name]** ("we", "us", or "our") respects your privacy and is committed to protecting your personal data. This Privacy Policy informs you about how we collect, process, and safeguard your personal data in accordance with the **General Data Protection Regulation (GDPR)** and the **California Consumer Privacy Act (CCPA)**.
* **Data Controller:** [Company Legal Name], [Corporate Address]
* **Data Protection Officer (DPO) Contact:** dpo@company.com

## 2. Categories of Personal Data We Collect
| Category of Data | Specific Data Elements Collected | Purpose of Collection |
|---|---|---|
| **Identity Data** | First name, last name, username, title | Account creation & user authentication |
| **Contact Data** | Email address, telephone number, billing address | Service notifications, billing, customer support |
| **Technical & Usage Data** | IP address, browser type, device identifiers, cookies | Security logging, error diagnostics, UX optimization |
| **User Content** | Text documents and prompts input into the application | AI processing strictly pursuant to user instructions |

## 3. Legal Grounds for Processing (GDPR Article 6)
We process personal data only when lawful bases apply:
* **Contractual Necessity:** Processing required to fulfill our Terms of Service and deliver software features.
* **Legitimate Interests:** Fraud prevention, platform cybersecurity, and debugging system errors.
* **Consent:** When you explicitly opt-in to marketing communications or optional browser cookies.

## 4. Third-Party Processors & Disclosures
We do not sell your personal data. We disclose data strictly to vetted sub-processors under Data Processing Agreements (DPAs):
* **Cloud Infrastructure:** Amazon Web Services (AWS) / Google Cloud Platform (Data hosting)
* **Payment Gateways:** Stripe Inc. (PCI-DSS compliant payment processing)
* **AI Model Providers:** OpenAI / Anthropic / Google (Subject to enterprise zero-data-retention APIs)

## 5. Cross-Border International Data Transfers
When transferring data outside the European Economic Area (EEA), we implement Standard Contractual Clauses (SCCs) approved by the European Commission, ensuring equivalent data protection.

## 6. Your Legal Rights as a Data Subject
Under applicable data protection laws, you possess the right to:
1. **Access & Portability:** Request copies of your personal data in a structured, machine-readable format.
2. **Rectification:** Correct inaccurate or incomplete personal records.
3. **Erasure ("Right to be Forgotten"):** Request permanent deletion of your personal data.
4. **Restriction & Objection:** Object to or restrict processing based on legitimate interests.
To exercise any of these rights, contact **privacy@company.com**. We respond within thirty (30) days without cost.

## 7. Data Retention & Security Architecture
We retain personal data only for as long as necessary to fulfill service purposes. All data in transit is encrypted using TLS 1.3, and data at rest is encrypted using AES-256 standards.`,
    aiPrompt: 'You are a certified Data Privacy Officer (CIPP/E) and regulatory compliance attorney. Draft an enterprise privacy policy adhering strictly to GDPR and CCPA requirements, complete with lawful processing bases, sub-processor tables, and data subject rights procedures.'
  },

  // ─── 📢 Marketing, PR & Communications ────────────────────────────────────
  'press-release': {
    id: 'press-release',
    title: '📰 Official AP-Style Press Release',
    category: 'marketing',
    badge: 'Marketing',
    desc: 'Associated Press standard news release: FOR IMMEDIATE RELEASE header, dateline, hook, executive quotes, proof metrics, and boilerplate.',
    sections: ['FOR IMMEDIATE RELEASE Header', 'Dateline & Lead Paragraph', 'Executive Leadership Quote', 'Key Innovations & Metrics Table', 'Partner / Customer Quote', 'Boilerplate & Media Contact'],
    skeleton: `FOR IMMEDIATE RELEASE

# [HEADLINE: Strong, Active Verb Announcing Major Newsworthy Breakthrough]
### [Sub-Headline: 1-2 Sentences Expanding on the Value Proposition and Strategic Context]

**[CITY, STATE/COUNTRY] — [Date, e.g. September 9, 2026]** — **[Company Name]**, a leader in [industry/domain], today announced [Major Breakthrough / Product Launch / Milestone]. [Follow with the 5 Ws: Who, What, When, Where, and Why this fundamentally shifts the competitive landscape].

"[Insert quote from Chief Executive Officer or Founder articulating the bold vision, the underlying market pain solved, and customer enthusiasm]," stated [Executive Name], [Title] at [Company Name]. "[Second sentence reinforcing strategic market momentum and future growth trajectory]."

## Key Innovations & Quantifiable Impact:
| Feature / Innovation | Technological Breakthrough | Quantified Customer Impact |
|---|---|---|
| **Autonomous Workflow Engine** | Zero-latency context processing | 80% reduction in document drafting turnaround |
| **Enterprise Privacy Core** | Local offline processing architecture | 100% confidential data governance compliance |
| **Universal Ecosystem Sync** | Multi-platform native Office integration | Instant deployment across 10,000+ enterprise seats |

"[Insert supporting quote from premier enterprise customer, strategic partner, or respected industry analyst validating real-world effectiveness]," added [Partner Name], [Partner Title] at [Partner Organization].

## Availability & Pricing
[State immediate availability dates, pricing tier overview, free trial access URL, or enterprise rollout schedule].

## About [Company Name]
[Company Name] is pioneering [one-sentence mission statement]. Founded in [Year] and headquartered in [City, State], the company empowers [target audience, e.g. millions of researchers, writers, and enterprise teams worldwide] with cutting-edge [technology category]. For more information, visit [https://www.company.com].

---

### Media Relations Contact:
* **Press Contact:** [Name of Communications Director]
* **Email:** press@company.com
* **Phone:** (555) 019-2834
* **Media Press Kit:** [https://www.company.com/press-kit]`,
    aiPrompt: 'You are a veteran corporate communications director and PR strategist. Write an AP-Style official press release featuring a scroll-stopping headline, compelling executive quotes, milestone impact table, and media boilerplate.'
  },

  'gtm-launch-strategy': {
    id: 'gtm-launch-strategy',
    title: '🚀 Go-To-Market (GTM) Product Launch Plan',
    category: 'marketing',
    badge: 'Marketing',
    desc: 'End-to-end commercial launch playbook: ICP buyer personas, positioning pillars, alpha/beta/GA rollout phases, and channel dashboard.',
    sections: ['Product Positioning Narrative', 'Ideal Customer Profile (ICP)', 'Pricing & Packaging Strategy', 'Phased Launch Execution Plan', 'Integrated Channel Strategy', 'Launch KPI Dashboard'],
    skeleton: `# Go-To-Market (GTM) Launch Strategy: [Product / Feature Name]
### Launch Target Date: [Date] | Product Marketing Lead: [PMM Name]

## 1. Product Positioning & Core Value Narrative
* **For:** [Target market / customer segment]
* **Who Struggle With:** [Specific high-friction operational pain point]
* **Our Product Is A:** [Product category description]
* **That Delivers:** [Primary transformative business benefit]
* **Unlike:** [Main legacy alternatives or primary competitor]
* **Our Unique Moat Is:** [Key proprietary differentiator and unfair advantage]

## 2. Ideal Customer Profile (ICP) & Buyer Personas
| Buyer Persona | Job Title & Department | Primary Pain Point | What Convinces Them to Buy |
|---|---|---|---|
| **Economic Buyer** | Chief Operating Officer / VP | Escalating overhead costs & slow turnaround | Hard ROI calculation (< 6 month payback) |
| **Technical Evaluator** | IT Security Director | Data leakage & compliance vulnerability | SOC-2 compliance & local processing options |
| **Daily End-User** | Knowledge Worker / Researcher | Tedious formatting & repetitive writing friction | Frictionless, intuitive 1-click workflows |

## 3. Pricing, Packaging & Commercial Tiers
* **Free / Starter Tier:** [Feature access hook designed for frictionless bottom-up product-led adoption]
* **Professional Tier ($X/month):** [Core power-user features, unlimited generations, priority models]
* **Enterprise Tier ($Y/seat/year):** [SSO, centralized billing, dedicated account manager, custom SLA]

## 4. Phased Launch Timeline & Milestones
| Phase | Duration | Target Audience | Primary Goal & Milestone |
|---|---|---|---|
| **Phase 1: Alpha (Private)** | Weeks 1–3 | 20 Trusted Design Partners | Identify bugs & refine onboarding flow |
| **Phase 2: Beta (Public)** | Weeks 4–6 | 500 Waitlist Community Users | Gather quantitative testimonials & NPS > 50 |
| **Phase 3: General Availability** | Week 7+ | Global Market | PR Blitz, ProductHunt launch, paid ad scale |

## 5. Integrated Marketing & Distribution Channel Strategy
* **Product Hunt & Tech Community:** [Launch day playbook, maker comment, founder video]
* **Content Marketing & SEO:** [5 high-intent comparison articles targeting competitor keywords]
* **Social Media Blitz:** [LinkedIn executive thought-leadership hooks and Twitter/X demo GIFs]
* **Direct Outbound Sales:** [Account-Based Marketing targeting top 200 high-fit enterprise accounts]

## 6. Launch Success KPI Dashboard
| Metric | 30-Day Target | 60-Day Target | 90-Day Target |
|---|---|---|---|
| New User Sign-ups | 5,000 | 15,000 | 40,000 |
| Free-to-Paid Conversion | 3.5% | 4.2% | 5.0% |
| Net New ARR Generated | $50,000 | $175,000 | $500,000 |`,
    aiPrompt: 'You are a Head of Product Marketing (PMM) and Go-To-Market consultant. Draft an end-to-end commercial GTM launch plan with positioning pillars, buyer persona matrix, phased rollout schedule, and 90-day KPI scorecard.'
  },

  'crisis-communication': {
    id: 'crisis-communication',
    title: '🚨 Executive Crisis Communication Plan',
    category: 'marketing',
    badge: 'Marketing',
    desc: 'Incident response playbook: Severity matrix, first-hour holding statements, stakeholder FAQ guides, and post-incident debrief.',
    sections: ['Severity Classification Matrix', 'Crisis Command Team Protocol', 'First-Hour Holding Statements', 'Stakeholder FAQ Playbook', 'Internal Employee Memo', 'Post-Mortem Review Protocol'],
    skeleton: `# Executive Crisis Communication & Incident Response Playbook
### Strict Confidentiality | Operational Protocol for [Company Name]

## 1. Incident Severity Classification Matrix
| Severity Tier | Incident Definition | Authorized Spokesperson | Response Time SLA |
|---|---|---|---|
| **Level 1 (Critical)** | Major data breach, critical system outage, legal investigation | Chief Executive Officer (CEO) | < 60 Minutes |
| **Level 2 (Moderate)** | Partial feature degradation, minor supplier disruption | VP of Communications / Operations | < 3 Hours |
| **Level 3 (Low)** | Isolated customer bug, minor social media complaint | Customer Support Lead | < 12 Hours |

## 2. Crisis Command Team & Decision Protocol
* **Incident Commander:** [CEO / COO Name] — Ultimate authority on external releases.
* **Legal Counsel:** [General Counsel Name] — Evaluates liability and regulatory disclosures.
* **Communications Lead:** [Head of PR Name] — Drafts and disseminates holding statements.
* **Technical Lead:** [CTO / CISO Name] — Establishes ground truth on remediation timeline.

## 3. First-Hour Holding Statements (Ready-to-Deploy)
### For Immediate Social Media / Status Page Publication:
> *"We are actively investigating an issue affecting [Service/Product Component]. Our engineering team is fully mobilized and working to resolve this as quickly as possible. We take this situation with the utmost seriousness and will share verified updates on this channel every 30 minutes. We deeply apologize for the inconvenience."*

### For Press / Inquiring Journalists:
> *"Earlier today, [Company Name] identified an operational disruption impacting [Service Name]. We immediately activated our incident response protocols to secure systems and restore full functionality. We are conducting a thorough investigation and will provide further details at [https://status.company.com]."*

## 4. Key Stakeholder Messaging & FAQ Playbook
* **Q: Was user personal data or financial information compromised?**
  * *Response:* "Based on our current investigation, [State verified facts only. Never speculate. State security protocols active]."
* **Q: When will regular operations be fully restored?**
  * *Response:* "Our teams have stabilized core operations, and we anticipate full recovery within [Timeline]. Updates are posted live on our status page."

## 5. Internal All-Hands Employee Alignment Memo
[Guidance instructing staff to direct all press inquiries to press@company.com and refrain from commenting on personal social channels]

## 6. Post-Mortem Review & Trust Restoration Protocol
* Publish transparent root-cause analysis (RCA) within 72 hours of incident resolution.
* Outline 3 concrete structural engineering or policy safeguards implemented to ensure non-recurrence.`,
    aiPrompt: 'You are an executive crisis communications specialist and corporate PR counsel. Write an emergency Incident Crisis Response Playbook with severity classification tables, holding statements, media FAQ playbooks, and post-mortem debrief procedures.'
  },

  // ─── 🎯 Career & Professional CV / Resume Studio ───────────────────────────
  'banking-finance-resume': {
    id: 'banking-finance-resume',
    title: '🏦 Banking & Financial Services CV / Resume',
    category: 'career',
    badge: 'Banking',
    desc: 'High-impact resume tailored for Commercial Banking, Credit Risk Underwriting, AML/KYC Compliance, Branch Management, and Financial Analysis.',
    sections: ['Candidate Header & Contact Info', 'Executive Banking Profile', 'Core Banking & Financial Competencies Matrix', 'Professional Banking Experience', 'Quantitative Lending & Risk Metrics', 'Education & Banking Certifications (CAMS, JAIBP, CFA)', 'Key Banking Achievements & Honors'],
    skeleton: `# [CANDIDATE FULL NAME]
**Senior Banking Operations & Credit Risk Specialist**
\`📧 [email@domain.com]\` • \`📱 [+1 (555) 019-2834]\` • \`📍 [City, State / Country]\` • \`🔗 [linkedin.com/in/profile]\`

---

## Executive Professional Profile
[Seasoned Banking & Financial Services Professional with **[X]+ years** of demonstrated leadership in commercial lending, credit risk underwriting, and branch operations across premier financial institutions. Proven track record managing **$[XX]M+** loan portfolios, reducing Non-Performing Loans (NPL) by **[X]%**, and maintaining **100% regulatory compliance** across State Bank / Federal Reserve / Basel III / AML-KYC mandates.]

---

## 🏦 Core Banking & Financial Competencies

| Banking Operations & Lending | Risk Management & Compliance | Financial Systems & Analytics |
| :--- | :--- | :--- |
| • Commercial & Retail Credit Appraisal | • AML / CFT & CDD Compliance (FATF) | • Core Banking (Temenos T24 / Finacle) |
| • Loan Syndication & SME Financing | • Basel III & IFRS 9 Risk Governance | • Advanced Financial Modeling (DCF / LBO) |
| • Treasury & Cash Management | • Credit Risk Grading & Internal Audit | • Bloomberg Terminal & Reuters Eikon |
| • Trade Finance (L/C, Guarantees) | • Non-Performing Loan (NPL) Remediation | • ERP Accounting (SAP FI-CO / Oracle) |

---

## 💼 Professional Banking Experience

### **[Senior Commercial Credit Officer / Branch Manager]** | [Premier Commercial Bank Name]
*[City, Country] • [Month, Year] – Present*
* Managed and expanded a diversified corporate and SME lending portfolio valued at **$[XX]M**, achieving **[XX]% YoY portfolio growth** while maintaining zero defaults.
* Spearheaded end-to-end credit appraisal, financial statement analysis, debt-service coverage (DSCR), and collateral valuation for **[XX]+ enterprise credit applications** annually.
* Enforced strict Anti-Money Laundering (AML), Counter-Terrorist Financing (CFT), and Know Your Customer (KYC) protocols, securing a **100% Satisfactory** rating in internal and regulatory audits.
* Restructured **$[X.X]M** in distressed credit facilities, reducing overall institutional Non-Performing Loan (NPL) ratio from **[X.X]% to [X.X]%** within 18 months.
* Mentored and led a cross-functional team of **[XX] relationship managers and credit analysts**, exceeding annual deposit acquisition targets by **[XX]% ($[XX]M)**.

### **[Credit Analyst / Relationship Associate]** | [National Investment / Commercial Bank]
*[City, Country] • [Month, Year] – [Month, Year]*
* Conducted rigorous ratio analysis, sensitivity stress-testing, and cash flow projections for corporate working capital lines and term finance facilities up to **$[XX]M**.
* Drafted high-stakes Credit Approval Memos (CAM) submitted directly to the Executive Credit Committee for multi-million dollar syndications.
* Automated credit monitoring workflows using core banking queries, accelerating credit file turnaround times by **[XX]%** without compromising risk tolerances.

---

## 🎓 Academic Credentials & Banking Certifications

| Degree / Professional Credential | Institution / Awarding Body | Year Completed |
| :--- | :--- | :--- |
| **Master of Science (MS) in Banking & Finance / MBA** | [University Name, City] | [Year] |
| **Certified Anti-Money Laundering Specialist (CAMS)** | ACAMS (Association of Certified AML Specialists) | [Year] |
| **Chartered Financial Analyst (CFA) / FRM** | CFA Institute / GARP | [Year] |
| **Junior / Senior Associate Banking Diploma (JAIBP / AIB)** | Institute of Bankers / Professional Banking Academy | [Year] |

---

## 🏆 Key Achievements & Banking Honors
* **Best Branch / Relationship Manager of the Year ([Year]):** Recognized across the [Region] network for highest asset growth (**+$[XX]M**) and lowest credit loss provision.
* **Special Regulatory Commendation:** Led seamless digital transition to Core Banking System (T24 / Finacle) with zero data discrepancy or operational downtime.`,
    aiPrompt: 'You are an executive banking recruiter and commercial banking talent partner. Write an elite, ATS-optimized Banking & Financial Services CV/Resume with high-impact Google XYZ metric bullets, Basel III / AML-KYC compliance, lending portfolio metrics, and core banking technical expertise.'
  },

  'executive-leadership-resume': {
    id: 'executive-leadership-resume',
    title: '💼 Executive Leadership & Management Resume',
    category: 'career',
    badge: 'Executive',
    desc: 'Boardroom-ready CV for C-Suite, VP, General Managers, and Management Consultants with P&L scale, M&A, and organizational transformation achievements.',
    sections: ['Executive Header', 'Value Proposition & Vision', 'Core Executive Competencies', 'Career Milestones & P&L Scale', 'Board Advisory & Governance', 'Education & Credentials'],
    skeleton: `# [EXECUTIVE FULL NAME]
**Chief Executive Officer / Chief Operating Officer / Vice President**
\`📧 [executive@domain.com]\` • \`📱 [+1 (555) 492-1102]\` • \`📍 [City, State / Country]\` • \`🔗 [linkedin.com/in/executive]\`

---

## Executive Value Proposition
[Transformational C-Suite Executive with **[XX]+ years** orchestrating sustainable top-line growth, operational excellence, and enterprise valuation expansion across multinational corporations. Architect of global go-to-market strategies and post-merger integrations yielding **$[XXX]M+ in enterprise value**.]

---

## ⚡ Core Executive Competencies

| Strategic Vision & P&L | Operational Excellence | Stakeholder Governance |
| :--- | :--- | :--- |
| • Global P&L Management ($[XXX]M+) | • Operational Scaling & LEAN Delivery | • Board of Directors Advisory & Reporting |
| • Mergers & Acquisitions (M&A) Due Diligence | • Supply Chain & Digital Automation | • Investor Relations & Venture Capital |
| • Enterprise Market Expansion & GTM | • Culture Transformation & Talent Retention | • Regulatory Compliance & Enterprise Risk |

---

## 💼 Executive Leadership Experience

### **[Chief Operating Officer / Vice President]** | [Multinational Enterprise Corporation]
*[City, Country] • [Month, Year] – Present*
* Stewarded enterprise operations and full P&L accountability for **$[XXX]M** annual operating budget across 5 global regions with **[X,XXX]+ personnel**.
* Delivered **[XX]% EBITDA margin expansion** in 24 months by restructuring procurement contracts and deploying automated cloud enterprise resource planning.
* Led operational integration of a **$[XX]M strategic acquisition**, capturing **$[XX]M in annualized synergies** 4 months ahead of schedule.
* Reduced voluntary high-performer turnover from **[XX]% to [X]%** by engineering a transparent merit-based equity compensation framework.

### **[Managing Director / Senior Vice President]** | [High-Growth Enterprise Firm]
*[City, Country] • [Month, Year] – [Month, Year]*
* Scaled business unit revenue from **$[XX]M to $[XXX]M** over 4 years through systematic geographic expansion and strategic tier-1 enterprise partnerships.
* Chaired Executive Risk and Capital Allocation Committees, sanctioning multi-million dollar R&D and capital expenditure initiatives.

---

## 🎓 Executive Education & Board Credentials

| Credential / Program | Institution | Year |
| :--- | :--- | :--- |
| **Master of Business Administration (MBA)** | [Prestigious Business School, e.g. Harvard / Wharton / INSEAD] | [Year] |
| **Certified Corporate Director (Governance)** | [Institute of Corporate Directors / NACD] | [Year] |
| **Bachelor of Science (B.Sc.) in Engineering / Economics** | [University Name] | [Year] |`,
    aiPrompt: 'You are an elite executive search partner at Spencer Stuart. Draft a boardroom-ready Executive Leadership Resume emphasizing multi-million dollar P&L scale, EBITDA margin expansion, M&A integration, and governance.'
  },

  'tech-software-resume': {
    id: 'tech-software-resume',
    title: '💻 Modern Tech & Software Engineer CV',
    category: 'career',
    badge: 'Tech',
    desc: 'Modern technical resume with two-column skills matrix (Languages, Cloud, System Design), high-scale architectural impact, and open-source contributions.',
    sections: ['Header & Socials', 'Technical Summary', 'Skills & Infrastructure Matrix', 'Software Engineering Experience', 'System Architecture Highlights', 'Education & Cloud Certifications'],
    skeleton: `# [DEVELOPER FULL NAME]
**Staff / Senior Full Stack Software Engineer & Cloud Architect**
\`📧 [dev@domain.com]\` • \`📱 [+1 (555) 782-9341]\` • \`📍 [City, State / Country]\` • \`💻 [github.com/profile]\` • \`🔗 [linkedin.com/in/profile]\`

---

## Technical Profile
[Forward-thinking Senior Software Engineer with **[X]+ years** architecting distributed, fault-tolerant cloud systems handling **[XX]M+ daily active requests**. Deep expertise across modern JavaScript/TypeScript, Python, Go, microservices, containerization, and enterprise cloud infrastructure (AWS/GCP).]

---

## 🛠️ Technical Skills & Infrastructure Matrix

| Category | Core Competencies & Technologies |
| :--- | :--- |
| **Programming Languages** | TypeScript, JavaScript (ESNext), Python, Go, SQL, HTML5/CSS3 |
| **Frameworks & Libraries** | React.js, Next.js, Node.js, Express, FastAPI, Django, TailwindCSS |
| **Databases & Caching** | PostgreSQL, MySQL, Redis, MongoDB, Elasticsearch, DynamoDB |
| **Cloud & DevOps (CI/CD)** | AWS (ECS, Lambda, S3, RDS), Docker, Kubernetes, Terraform, GitHub Actions |
| **Architecture & Testing** | Microservices, Event-Driven (Kafka/RabbitMQ), RESTful, GraphQL, Jest, Cypress |

---

## 💼 Software Engineering Experience

### **[Senior Full Stack Engineer / Tech Lead]** | [Leading Tech Enterprise / SaaS Startup]
*[City, Country] • [Month, Year] – Present*
* Architected and delivered a distributed real-time messaging pipeline handling **[XX]K events/sec**, reducing API response latency by **[XX]% (from 420ms to 85ms)**.
* Directed migration from legacy monolithic architecture to containerized microservices on AWS EKS, resulting in a **[XX]% reduction in monthly AWS cloud expenditure ($[XX]K/yr)**.
* Spearheaded full-stack engineering for core customer portal serving **[XXX]K+ active users**, boosting user session retention by **[XX]%**.
* Instituted automated CI/CD deployment gates and test suites (Jest/Cypress), elevating overall unit test coverage from **[XX]% to [XX]%**.

### **[Software Engineer]** | [Technology Solutions Firm]
*[City, Country] • [Month, Year] – [Month, Year]*
* Developed high-throughput RESTful and GraphQL APIs in Node.js and PostgreSQL, achieving **99.99% uptime SLA**.
* Designed caching tier with Redis, decreasing database read pressure by **[XX]%** during high-concurrency peak traffic periods.

---

## 🎓 Education & Cloud Certifications

| Certification / Degree | Issuing Organization | Year |
| :--- | :--- | :--- |
| **AWS Certified Solutions Architect – Professional** | Amazon Web Services | [Year] |
| **Certified Kubernetes Administrator (CKA)** | Linux Foundation / CNCF | [Year] |
| **B.S. in Computer Science / Software Engineering** | [University Name, City] | [Year] |`,
    aiPrompt: 'You are a Principal Software Engineer and tech hiring manager at Google. Write an elite Software Engineer Resume with quantified Google XYZ metric bullets (latency drops, scale, AWS cost savings), two-column skills matrix, and clean microservices engineering experience.'
  },

  'academic-scholar-cv': {
    id: 'academic-scholar-cv',
    title: '🎓 Academic Scholar & Faculty Curriculum Vitae',
    category: 'career',
    badge: 'Academic',
    desc: 'Comprehensive academic curriculum vitae with research interests, peer-reviewed publications, research grants, teaching pedagogy, and international conference presentations.',
    sections: ['Academic Header', 'Appointments & Education', 'Research Interests', 'Peer-Reviewed Publications', 'Funded Grants & Awards', 'Teaching Experience & Advising', 'Academic Service & References'],
    skeleton: `# [SCHOLAR FULL NAME, Ph.D.]
**Associate Professor of [Field / Department]**
\`📧 [professor@university.edu]\` • \`📱 [+1 (555) 234-8891]\` • \`🏛️ [Department of X, University Name, City, Country]\` • \`🌐 [scholar.google.com/citations?user=xyz]\`

---

## Academic Appointments & Affiliations
* **Associate Professor (Tenured):** Department of [Subject], [University Name] (*[Year] – Present*)
* **Assistant Professor:** Department of [Subject], [University Name] (*[Year] – [Year]*)
* **Postdoctoral Research Fellow:** [Laboratory / Institute Name], [University Name] (*[Year] – [Year]*)

---

## 🎓 Education
* **Ph.D. in [Field / Subject]:** [University Name], [Year]
  * *Dissertation:* "[Title of Doctoral Dissertation]"
  * *Advisors:* Prof. [Advisor 1], Prof. [Advisor 2]
* **M.Sc. in [Field / Subject]:** [University Name], [Year] (*First Class Honors*)
* **B.Sc. in [Field / Subject]:** [University Name], [Year] (*Summa Cum Laude*)

---

## 🔬 Primary Research Interests
[Theoretical and empirical exploration of **[Topic 1]**, **[Topic 2]**, and **[Topic 3]**, with specific focus on quantitative computational modeling and experimental validation.]

---

## 📚 Peer-Reviewed Journal Publications
1. **[Candidate Name]**, [Co-Author 1], & [Co-Author 2] ([Year]). "[Title of Landmark Paper]". *Journal of Premier Scientific Research*, **[Vol]**([Issue]), [pp. XX–XX]. https://doi.org/10.xxxx/xxxx
2. **[Candidate Name]** & [Co-Author] ([Year]). "[Title of Second Major Publication]". *International Review of [Discipline]*, **[Vol]**([Issue]), [pp. XX–XX].
3. [Co-Author] & **[Candidate Name]** ([Year]). "[Title of Collaborative Empirical Study]". *Annals of [Discipline]*, **[Vol]**([Issue]), [pp. XX–XX].

---

## 💰 Funded Research Grants & Fellowships

| Grant Title / Number | Funding Agency | Amount Awarded | Role & Tenure |
| :--- | :--- | :--- | :--- |
| *[Grant Name / Title of Research Project]* | National Science Foundation (NSF) | **$[XXX],000** | Principal Investigator (PI), [Year–Year] |
| *[Interdisciplinary Collaborative Grant]* | European Research Council (ERC) | **€[XXX],000** | Co-Investigator (Co-PI), [Year–Year] |

---

## 👨‍🏫 University Teaching & Doctoral Mentorship
* **Courses Taught:** Advanced [Course 1] (Graduate), Foundations of [Course 2] (Undergraduate).
* **Doctoral Students Supervised:** Advised **[X] Ph.D. dissertations** and **[XX] Master's theses** to successful defense.`,
    aiPrompt: 'You are an Ivy League academic dean. Write an elite, publication-rich Academic Scholar Curriculum Vitae (CV) with appointments, doctoral credentials, peer-reviewed journal articles in APA format, NSF/ERC research grants, and university teaching experience.'
  },

  'teaching-educator-resume': {
    id: 'teaching-educator-resume',
    title: '👩‍🏫 Teaching & Education Professional CV / Resume',
    category: 'career',
    badge: 'Education',
    desc: 'Comprehensive CV for School Teachers, Lecturers, STEM Educators, and Department Chairs with pedagogy, curriculum design, student outcomes, and EdTech.',
    sections: ['Educator Header & Contact Info', 'Teaching Philosophy & Profile', 'Core Educational Competencies Matrix', 'Teaching & Academic Leadership Experience', 'Standardized Outcomes & Student Achievements', 'Education & Teaching Certifications', 'Professional Affiliations & Workshops'],
    skeleton: `# [EDUCATOR FULL NAME, M.Ed. / B.Ed.]
**Senior STEM / Secondary Education Specialist & Curriculum Developer**
\`📧 [educator@school.edu]\` • \`📱 [+1 (555) 381-9042]\` • \`📍 [City, State / Country]\` • \`🔗 [linkedin.com/in/educator]\`

---

## Teaching Philosophy & Professional Profile
[Dedicated, student-centered educator with **[X]+ years** of classroom excellence fostering critical inquiry, academic rigor, and inclusive learning environments. Proven record designing differentiated curriculum frameworks, integrating educational technologies, and elevating standardized student assessment pass rates by **[XX]%** across diverse cohorts.]

---

## 📚 Core Educational Competencies & Pedagogical Expertise

| Curriculum & Instructional Design | Classroom Leadership & Assessment | Educational Technology & LMS |
| :--- | :--- | :--- |
| • Differentiated Learning (IEP/504) | • Formative & Summative Evaluation | • Google Classroom & Canvas LMS |
| • Backward Design (UbD Framework) | • Positive Behavioral Support (PBIS) | • Interactive Smartboard & EdTech Apps |
| • Inquiry-Based STEM / STEAM Learning | • Rubric & Standardized Test Prep | • Student Performance Analytics |
| • Cross-Disciplinary Project Delivery | • Parent-Teacher Communication Strategy | • Digital Hybrid & Blended Learning |

---

## 💼 Teaching & Academic Leadership Experience

### **[Lead High School / Middle School Teacher]** | [School / College District Name]
*[City, Country] • [Month, Year] – Present*
* Delivered rigorous, standards-aligned instruction across **[Subject / Grade Level]** for **[XXX]+ students** annually, maintaining an average student satisfaction rating of **[XX]%**.
* Advanced student mastery on state/national standardized assessments, raising cohort proficiency rates from **[XX]% to [XX]%** within two academic cycles.
* Engineered and implemented **[X] innovative curriculum units** incorporating project-based learning and digital simulation tools, boosting classroom assignment completion by **[XX]%**.
* Mentored **[X] student teachers and junior faculty members**, facilitating quarterly professional development seminars on data-driven formative assessment.
* Partnered proactively with parents, counselors, and administrative staff to craft personalized academic intervention plans for **[XX]+ at-risk students**, recovering **[XX]%** to grade-level benchmark.

### **[Subject Teacher / Educational Associate]** | [Academy / Preparatory School Name]
*[City, Country] • [Month, Year] – [Month, Year]*
* Taught foundational and honors curricula to classes of **[XX–XX] students**, fostering collaborative discussion and analytical writing skills.
* Integrated digital gamification and adaptive learning software, increasing student homework submission consistency by **[XX]%**.

---

## 🎓 Academic Credentials & Teaching Licensure

| Degree / Teaching Credential | Awarding Institution | Year Completed |
| :--- | :--- | :--- |
| **Master of Education (M.Ed.) in Curriculum & Instruction** | [University Name, City] | [Year] |
| **State / National Professional Educator License ([Subject / Grades])** | [Department of Education] | [Valid through Year] |
| **Advanced Placement (AP) / IB Certified Educator** | College Board / International Baccalaureate | [Year] |
| **Bachelor of Science / Arts (B.S./B.A.) in [Subject Area]** | [University Name, City] | [Year] |

---

## 🏆 Honors, Grants & Committee Leadership
* **Excellence in Teaching Award ([Year]):** Selected by peers and administration for transformative classroom impact and student mentorship.
* **School District Innovation Grant ($[X],000):** Secured competitive district grant to establish a collaborative robotics / digital media laboratory.`,
    aiPrompt: 'You are an education recruiter and school district superintendent. Write an inspiring, highly professional Teaching & Education CV/Resume emphasizing differentiated pedagogy, student test score gains, curriculum design, classroom leadership, and certified licensure credentials.'
  },

  'medical-healthcare-resume': {
    id: 'medical-healthcare-resume',
    title: '🏥 Healthcare, Nursing & Clinical Specialist CV',
    category: 'career',
    badge: 'Healthcare',
    desc: 'Clinical CV for Doctors, Registered Nurses (RN/BSN), Healthcare Administrators, and Clinical Specialists with triage, patient outcomes, EHR systems, and HIPAA.',
    sections: ['Clinical Header & Licensure', 'Clinical Practice Summary', 'Medical Competencies & Systems Matrix', 'Clinical Healthcare Experience', 'Patient Care & Quality Metrics', 'Education & Board Certifications', 'Life Support & Professional Credentials'],
    skeleton: `# [CLINICAL PROFESSIONAL NAME, RN / BSN / MD]
**Senior Clinical Nurse Specialist & Acute Patient Care Coordinator**
\`📧 [nurse@hospital.org]\` • \`📱 [+1 (555) 712-4490]\` • \`📍 [City, State / Country]\` • \`🏥 [State License: RN-XXXXX]\`

---

## Clinical Practice Summary
[Compassionate, board-certified clinical healthcare professional with **[X]+ years** of acute care, triage, and patient advocacy leadership across high-volume hospital trauma centers and intensive care units. Expertise administering advanced medical interventions, maintaining **100% Joint Commission / HIPAA compliance**, and mentoring nursing staff to optimize patient safety outcomes.]

---

## 🩺 Clinical Competencies & Hospital Systems

| Patient Care & Acute Nursing | Clinical Governance & Safety | Medical Systems & EHR |
| :--- | :--- | :--- |
| • Emergency Triage & Trauma Resuscitation | • Joint Commission (JCAHO) Protocols | • Epic Systems (Inpatient / Ambulatory) |
| • Hemodynamic & Vital Signs Monitoring | • Medication Reconciliation & Pyxis | • Cerner Millennium & Meditech |
| • Post-Operative Critical Care Recovery | • Infection Control & Sterile Technique | • Automated IV Infusion & Telemetry |
| • Patient & Family Discharge Education | • HIPAA Patient Confidentiality Safeguards | • Point-of-Care Diagnostic Testing |

---

## 💼 Clinical Healthcare Experience

### **[Charge Nurse / Clinical Care Specialist]** | [Metropolitan General Hospital]
*[City, Country] • [Month, Year] – Present*
* Coordinated holistic bedside nursing care for **[XX] acute-care patient beds** in a high-acuity medical-surgical / ICU ward, achieving zero medication administration errors.
* Reduced 30-day patient readmission rates by **[XX]%** through structured multi-lingual patient education and post-discharge continuity protocols.
* Supervised and directed daily shift assignments for **[XX] registered nurses, licensed practical nurses, and certified nursing assistants**.
* Participated as a core member of the Hospital Quality Improvement Committee, co-authoring infection reduction protocols that decreased central line infections by **[XX]%**.

### **[Staff Registered Nurse]** | [Regional Healthcare Center]
*[City, Country] • [Month, Year] – [Month, Year]*
* Conducted rapid triage assessments, stabilized incoming trauma cases, and administered specialized IV pharmacotherapy under physician direction.
* Documented comprehensive patient health histories and clinical assessments in Epic EHR with **100% timeliness and chart audit compliance**.

---

## 🎓 Medical Degrees & Professional Licensure

| Licensure / Degree | Issuing Board / Institution | License # / Year |
| :--- | :--- | :--- |
| **Registered Nurse (RN) Multi-State Compact License** | State Board of Nursing | License #RN-[XXXXX] |
| **Bachelor of Science in Nursing (BSN) / MSN** | [University School of Nursing, City] | [Year] |
| **Basic Life Support (BLS) & Advanced Cardiac Life Support (ACLS)** | American Heart Association (AHA) | [Current / Renewal Year] |
| **Pediatric Advanced Life Support (PALS)** | American Heart Association (AHA) | [Current / Renewal Year] |`,
    aiPrompt: 'You are a Chief Nursing Officer and healthcare hiring director. Write a compassionate, clinical, and precise Healthcare & Registered Nurse Resume highlighting patient safety metrics, EHR charting (Epic), acute trauma triage, and clinical certifications.'
  },

  'accounting-finance-resume': {
    id: 'accounting-finance-resume',
    title: '📊 Accounting, Audit & Corporate Finance Resume',
    category: 'career',
    badge: 'Accounting',
    desc: 'Executive resume for Certified Public Accountants (CPA/ACCA), Financial Controllers, Internal Auditors, and Tax Managers with GAAP/IFRS, SOX, and ERP.',
    sections: ['Accounting Header', 'Executive Financial Summary', 'Core Accounting & Audit Competencies', 'Corporate Accounting Experience', 'Audit Readiness & SOX Metrics', 'Credentials (CPA, ACCA, CMA)', 'Education & Degrees'],
    skeleton: `# [FINANCE PROFESSIONAL NAME, CPA / ACCA]
**Senior Corporate Controller & Financial Audit Specialist**
\`📧 [cpa@finance.com]\` • \`📱 [+1 (555) 832-1920]\` • \`📍 [City, State / Country]\` • \`🔗 [linkedin.com/in/cpa-profile]\`

---

## Executive Financial Profile
[Strategic Certified Public Accountant (CPA) with **[X]+ years** leading corporate accounting, general ledger stewardship, financial statement reporting, and SOX internal control compliance for multinational enterprises. Championed enterprise ERP migrations, accelerated monthly financial closes from **10 to 4 days**, and saved **$[X.X]M** through automated audit workflows.]

---

## 📊 Core Accounting, Tax & Audit Competencies

| Financial Accounting & Reporting | Audit, Risk & Compliance | ERP & Financial Systems |
| :--- | :--- | :--- |
| • US GAAP & IFRS Technical Accounting | • SOX 404 Internal Controls & Testing | • SAP S/4HANA (FI/CO Modules) |
| • Month-End & Year-End Close Acceleration | • Big 4 External Audit Management | • Oracle NetSuite & Hyperion HFM |
| • Balance Sheet Reconciliation & Accruals | • Corporate Tax Planning & Provisions | • Advanced Excel (VBA, Power Query, Macros) |
| • Financial Statement Consolidation | • Cost Accounting & Variance Analysis | • BlackLine Account Reconciliation |

---

## 💼 Corporate Accounting Experience

### **[Corporate Financial Controller / Accounting Manager]** | [Global Enterprise Corporation]
*[City, Country] • [Month, Year] – Present*
* Supervised end-to-end accounting operations and financial statement consolidations for **[X] international subsidiaries** generating **$[XXX]M in combined revenue**.
* Streamlined the monthly financial close cadence, reducing turnaround duration from **10 business days to 4 business days** through automated ERP journal entries.
* Managed relationships with external Big 4 auditing firms (PwC / EY / Deloitte / KPMG), achieving **100% clean audit opinions with zero material weaknesses** for 4 consecutive fiscal years.
* Spearheaded corporate internal control overhaul under SOX 404 guidelines, mitigating enterprise risk exposure across **[XX]+ core business operational cycles**.

### **[Senior Financial Auditor / Senior Accountant]** | [Public Accounting / Corporate Firm]
*[City, Country] • [Month, Year] – [Month, Year]*
* Planned and conducted comprehensive substantive testing, revenue recognition audits, and internal control reviews for commercial clients with revenues up to **$[XX]M**.
* Identified and resolved complex GAAP discrepancies in lease accounting (ASC 842) and revenue contracts (ASC 606), preventing multi-million dollar reporting adjustments.

---

## 🎓 Professional Designations & Education

| Certification / Degree | Awarding Organization | Year |
| :--- | :--- | :--- |
| **Certified Public Accountant (CPA)** | State Board of Accountancy | [Year] |
| **Association of Chartered Certified Accountants (ACCA) / CMA** | Global Accounting Body | [Year] |
| **Master of Accounting (MAcc) / B.S. in Accounting** | [University Name, City] | [Year] |`,
    aiPrompt: 'You are an Audit Partner at a Big 4 accounting firm. Write an authoritative, meticulous Accounting & Corporate Controller Resume focusing on US GAAP/IFRS technical reporting, SOX 404 compliance, Big 4 audit readiness, and month-end close acceleration.'
  },

  'sales-marketing-resume': {
    id: 'sales-marketing-resume',
    title: '📢 Sales, Digital Marketing & Growth Resume',
    category: 'career',
    badge: 'Growth & Sales',
    desc: 'High-converting resume for Enterprise Account Executives, VP of Sales, Digital Marketing Directors, and Growth Marketers with quota metrics, ARR, and ROI.',
    sections: ['Revenue Leader Header', 'Commercial Growth Profile', 'Go-To-Market & Revenue Skills Table', 'Sales & Marketing Experience', 'Quantified Quota & ARR Achievements', 'Marketing Tech Stack & CRM', 'Education'],
    skeleton: `# [COMMERCIAL LEADER FULL NAME]
**Vice President of Sales & Global Revenue Growth**
\`📧 [sales@growth.com]\` • \`📱 [+1 (555) 604-3319]\` • \`📍 [City, State / Country]\` • \`🔗 [linkedin.com/in/sales-leader]\`

---

## Commercial Growth & Revenue Profile
[High-impact revenue generation leader with **[X]+ years** scaling enterprise B2B sales pipelines, architecting full-funnel digital marketing engines, and exceeding commercial targets across competitive technology and SaaS markets. Track record delivering **$[XX]M+ in net-new ARR**, maintaining **[XXX]% quota attainment**, and expanding enterprise contract values by **[XX]%**.]

---

## 📈 Sales, Marketing & GTM Competencies

| Enterprise B2B Sales & Closing | Digital Marketing & Demand Gen | Sales Tech & CRM Platforms |
| :--- | :--- | :--- |
| • Enterprise SaaS Contract Closing ($[X]M+) | • Multi-Channel Paid Ads (Google, LinkedIn) | • Salesforce CRM & HubSpot Enterprise |
| • Outbound Account-Based Selling (ABM) | • Inbound SEO & Content Demand Generation | • Outreach.io, Gong.io, Salesloft |
| • MEDDPICC & Challenger Sales Methodology | • Customer Acquisition Cost (CAC) Optimization | • Google Analytics 4 & Mixpanel |
| • High-Velocity Inside Sales Team Building | • Product-Led Growth (PLG) & Conversion Rate | • ZoomInfo & Apollo Lead Generation |

---

## 💼 Professional Sales & Marketing Experience

### **[Director of Enterprise Sales / Head of Growth]** | [High-Growth SaaS Corporation]
*[City, Country] • [Month, Year] – Present*
* Generated **$[XX]M in net-new enterprise ARR** over 24 months, achieving **[XXX]% of assigned annual quota** and ranking in the top 1% of the global commercial organization.
* Recruited, trained, and led a high-performing team of **[XX] Account Executives and SDRs**, accelerating median team quota attainment from **[XX]% to [XX]%**.
* Closed company's largest flagship deal in history (**$[X.X]M multi-year contract**) through rigorous MEDDPICC stakeholder alignment and C-level executive sponsorship.
* Redesigned digital customer acquisition funnel, reducing Customer Acquisition Cost (CAC) by **[XX]%** while increasing marketing-qualified leads (MQLs) by **[XX]%**.

### **[Senior Enterprise Account Executive]** | [B2B Technology Solutions Provider]
*[City, Country] • [Month, Year] – [Month, Year]*
* Managed end-to-end sales cycles from prospecting through contract negotiation for Fortune 500 accounts, maintaining an average win rate of **[XX]%**.
* Expanded existing customer annual contract value (ACV) by **[XX]%** through strategic cross-sell and upsell initiatives.

---

## 🎓 Education & Professional Accreditations
* **Bachelor of Business Administration (BBA) in Marketing / Management:** [University Name], [Year]
* **Certified Enterprise Sales Professional (MEDDPICC Accredited):** [Year]
* **HubSpot Inbound & Google Ads Certified:** [Year]`,
    aiPrompt: 'You are an executive revenue recruiter and Chief Commercial Officer. Write a high-converting Sales & Digital Marketing Resume packed with aggressive quota attainment percentages (140%+), net-new ARR numbers, CAC reductions, and enterprise deal closures.'
  },

  'project-product-resume': {
    id: 'project-product-resume',
    title: '🛠️ Product & Agile Project Manager Resume',
    category: 'career',
    badge: 'Product & Agile',
    desc: 'Modern resume for Technical Product Managers, Scrum Masters, Agile Project Leads, and PMP professionals with sprint velocity, roadmap delivery, and user growth.',
    sections: ['Product Leader Header', 'Product Strategy Summary', 'Product Management & Agile Matrix', 'Product Leadership Experience', 'Feature Launch & Adoption Metrics', 'Certifications (PMP, CSM, ACP)', 'Education'],
    skeleton: `# [PRODUCT MANAGER FULL NAME, PMP / CSM]
**Senior Technical Product Manager & Agile Delivery Lead**
\`📧 [pm@product.io]\` • \`📱 [+1 (555) 921-7783]\` • \`📍 [City, State / Country]\` • \`🔗 [linkedin.com/in/product-lead]\`

---

## Product Leadership & Strategy Profile
[Visionary Senior Product Manager with **[X]+ years** orchestrating the discovery, engineering delivery, and commercial scaling of cloud and mobile applications used by **[X]M+ monthly active users**. Expert at translating ambiguous user friction into prioritized product roadmaps, leading cross-functional engineering squads, and optimizing core retention KPIs.]

---

## 🛠️ Product Management & Delivery Competencies

| Product Strategy & Discovery | Agile Execution & Engineering Delivery | Analytics, Data & UX |
| :--- | :--- | :--- |
| • Product Vision & Strategic Roadmapping | • Scrum, Kanban & Scaled Agile (SAFe) | • Product Analytics (Mixpanel, Amplitude) |
| • User Journey Mapping & Persona Research | • Sprint Planning, Backlog Refinement | • A/B Testing & Multivariate Experiments |
| • Customer Problem Discovery & Interviews | • Jira, Confluence, Linear, Asana | • Wireframing & Prototyping (Figma) |
| • Feature Prioritization (RICE / MoSCoW) | • Technical Spec (PRD) Authoring | • SQL Data Extraction & Funnel Analysis |

---

## 💼 Product Leadership Experience

### **[Lead Product Manager / Technical PM]** | [Leading Software & Technology Firm]
*[City, Country] • [Month, Year] – Present*
* Spearheaded end-to-end product strategy and delivery for flagship enterprise platform serving **[XXX]K+ daily enterprise users**, boosting Daily Active User (DAU) retention by **[XX]%**.
* Authored **[XX]+ Product Requirement Documents (PRDs)** and user stories for a cross-functional squad of **[XX] software engineers, UX designers, and QA leads**.
* Launched real-time collaboration feature that reduced user workflow drop-off by **[XX]%** and drove **$[X.X]M in annualized incremental upsell revenue**.
* Increased engineering team sprint velocity by **[XX]%** by refining acceptance criteria and eliminating scope creep during sprint planning cycles.

### **[Agile Project Manager / Scrum Master]** | [Digital Transformation Consultancy]
*[City, Country] • [Month, Year] – [Month, Year]*
* Managed simultaneous delivery of 4 multi-platform digital engineering initiatives on time and **[XX]% under allocated client budget**.
* Facilitated daily standups, sprint retrospectives, and backlog grooming sessions, improving team story point delivery predictability to **95%**.

---

## 🎓 Product Accreditations & Education

| Credential / Degree | Issuing Organization | Year |
| :--- | :--- | :--- |
| **Project Management Professional (PMP)** | Project Management Institute (PMI) | [Year] |
| **Certified Scrum Product Owner (CSPO) / CSM** | Scrum Alliance | [Year] |
| **B.S. in Information Systems / Software Engineering** | [University Name, City] | [Year] |`,
    aiPrompt: 'You are a VP of Product at a Silicon Valley tech unicorn. Write a data-driven Product & Agile Project Manager Resume highlighting user adoption metrics, RICE feature prioritization, sprint velocity improvements, and PMP/CSM certifications.'
  },

  'legal-counsel-resume': {
    id: 'legal-counsel-resume',
    title: '⚖️ Legal Counsel & Corporate Attorney CV',
    category: 'career',
    badge: 'Legal Counsel',
    desc: 'Distinguished CV for Corporate Attorneys, In-House Counsel, Commercial Litigators, and Compliance Directors with deal volume, contracts, and regulatory governance.',
    sections: ['Attorney Header & Bar Admissions', 'Executive Legal Practice Summary', 'Core Legal Competencies & Practice Areas', 'Legal Practice Experience', 'High-Stakes Transactional & Litigation Milestones', 'Law Degrees (J.D. / LL.M.)', 'Bar Admissions & Memberships'],
    skeleton: `# [ATTORNEY FULL NAME, ESQ. / LL.M.]
**Senior Corporate Counsel & Regulatory Affairs Attorney**
\`📧 [attorney@law.com]\` • \`📱 [+1 (555) 512-8874]\` • \`📍 [City, State / Country]\` • \`⚖️ [Bar Admission: Admitted in State/Court]\`

---

## Executive Legal Practice Summary
[Distinguished corporate counsel with **[X]+ years** of high-stakes transactional, contract negotiation, and enterprise risk management experience representing Fortune 500 corporations and private equity sponsors. Negotiated and closed **$[XXX]M+ in complex commercial transactions**, managed outside litigation counsel, and ensured full regulatory alignment with SEC / FTC / GDPR / CCPA standards.]

---

## ⚖️ Legal Competencies & Areas of Practice

| Corporate & Commercial Transactions | Regulatory Compliance & Governance | Intellectual Property & Litigation |
| :--- | :--- | :--- |
| • Commercial Contracts Drafting (MSA, SOW) | • Regulatory Affairs (SEC, FTC, DOJ) | • Trademark, Copyright & IP Licensing |
| • Mergers & Acquisitions (M&A) Due Diligence | • Global Privacy & Data Governance (GDPR) | • Commercial Litigation Management |
| • Strategic Joint Ventures & Partnership Deals | • Corporate Board Governance & Resolutions | • Employment Law & Executive Compensation |
| • Cross-Border Vendor & Supply Agreements | • Internal Ethics Investigations & Audits | • Alternative Dispute Resolution (ADR) |

---

## 💼 Legal Practice Experience

### **[Senior Corporate Counsel / Legal Director]** | [Global Corporation / Premier Law Firm]
*[City, Country] • [Month, Year] – Present*
* Drafted, negotiated, and successfully executed **[XXX]+ high-value commercial agreements** annually valued in excess of **$[XXX]M**, minimizing enterprise liability exposure.
* Guided executive leadership through **$[XX]M M&A transaction**, conducting thorough legal due diligence, antitrust review, and definitive acquisition agreement negotiation.
* Managed enterprise-wide data privacy and cyber governance compliance across 14 international jurisdictions under GDPR and state statutory frameworks.
* Supervised outside litigation counsel across active commercial disputes, resolving **[XX]% of proceedings through favorable pre-trial settlement** and saving **$[X.X]M in legal fees**.

### **[Associate Corporate Attorney]** | [Top-Tier National Law Firm]
*[City, Country] • [Month, Year] – [Month, Year]*
* Drafted governance documents, venture financing agreements, and SEC regulatory filings for institutional clients.
* Conducted in-depth legal research and authored persuasive litigation briefs submitted before state and federal appellate courts.

---

## 🎓 Legal Education & Bar Admissions

| Credential / Degree | Institution / Jurisdiction | Year |
| :--- | :--- | :--- |
| **Juris Doctor (J.D.) / Master of Laws (LL.M.)** | [Law School / University Name, City] | [Year] |
| **State Bar Admission (Active License in Good Standing)** | Supreme Court of [State / National Bar Association] | [Bar #XXXXX, Year] |
| **B.A. in Political Science / Economics / Philosophy** | [University Name, City] | [Summa Cum Laude, Year] |`,
    aiPrompt: 'You are a General Counsel at a Fortune 100 enterprise. Write an authoritative Legal Counsel & Corporate Attorney CV highlighting multi-million dollar contract negotiations, M&A due diligence, regulatory compliance (SEC, GDPR), and active state bar credentials.'
  }
};

if (typeof module !== 'undefined' && module.exports) {
  module.exports = { FRAMEWORK_DEFINITIONS };
}
