"""
Content library for the presales-deck-generator skill.

This is the single source of truth for the "static" parts of every pre-sales
deck: general ELCA facts, the Data & AI Business Unit content, the reusable
reference-project library, and industry modules. build_presales_deck.py
imports this file directly.

WHO EDITS THIS FILE: whoever ELCA's Data & AI BU nominates as skill
champions (see the department's shared plugin marketplace). No engineering
background is required — it's plain Python dicts and lists, organised by
clearly labelled section. Keep entries short: this content is meant to be
read off a slide, not a report.

Source of the current content: the 2026 ELCA Corporate Deck and the
Data, Analytics & AI Business Line intro slides Ankit Gupta's team already
reuses on client engagements (both on SharePoint under
PublicCorporateAssets/Corp Decks PPTs/2026). Re-pull from there whenever
those decks are refreshed — this file will drift out of date otherwise.
"""

# ---------------------------------------------------------------------------
# SECTION 1 — General ELCA content (About ELCA)
# ---------------------------------------------------------------------------

GENERAL_ELCA = {
    "glance": {
        "title": "ELCA at a Glance",
        "subtitle": "Key facts about the Group",
        "pillars": [
            ("1968", ['Founded as "electro-calcul" to build the software control of the Grande Dixence Dam']),
            ("2'350+", ["Headcount — mostly highly skilled IT professionals with university degrees"]),
            ("CHF 339M", ["Turnover"]),
            ("55+ years", ["Experience — ISO 9001, 14001, 27001, 27017, 27018 & 27701 certified"]),
        ],
    },
    "local_partner": {
        "title": "Your Local Partner, Globally Delivered",
        "subtitle": "Proximity where it matters, scale where it helps",
        "pillars": [
            ("Switzerland", [
                "1'380 people locally",
                "Pully (HQ), Geneva, Bern, Zurich, Basel, Rapperswil",
                "Proximity to clients for efficient collaboration",
            ]),
            ("Shoring Platform", [
                "2'390 people across our shoring platform",
                "Nearshore: Spain, Italy",
                "Offshore: Mauritius, Vietnam",
                "Strategic resilience, competitive advantage, access to diverse talent",
            ]),
        ],
    },
    "group_structure": {
        "title": "A Group Built From Entrepreneurial Units",
        "subtitle": "A set of entrepreneurial units within a unified Group",
        "pillars": [
            ("Industry Business Lines", [
                "Public administration", "Defense", "Health",
                "Transport", "Financial Services", "Retail & Industries",
            ]),
            ("Delivery Business Lines", [
                "Data, Analytics & AI", "Microsoft Solutions",
                "Java Engineering · .NET Engineering", "Architecture",
                "Digital Agency · ELCA Advisory", "Viacar · Swiss Business Solutions",
            ]),
            ("Shared Services & Shoring", [
                "HR · Finance · Corp Dev · Marketing & Corp Comm",
                "IT · Security & Privacy · Legal · Sales",
                "",
                "International Shoring Centers: Italy, Spain, Mauritius, Vietnam",
            ]),
        ],
    },
    "services": {
        "title": "We Offer a Broad Range of Services",
        "subtitle": "ELCA's in-house solutions and products, end to end",
        "pillars": [
            ("Engineering", ["Application & Mobile, UX, Web, Front-End, Interface & API, Cybersecurity"]),
            ("Integration", ["ERP, CRM, IIM Services & Solutions, Collaboration, CMS, BI, AI"]),
            ("Operations", ["Managed Services, Cloud Engineering, SaaS, CI, Application Management"]),
            ("Consulting", ["Digital Transformation, IT Strategy, IT-Management, Business Analysis"]),
        ],
    },
}

# ---------------------------------------------------------------------------
# SECTION 2 — Data, Analytics & AI Business Unit content
# ---------------------------------------------------------------------------

DATAAI_BU = {
    "mission": {
        "title": "We Love to Make You Successful With Data, Analytics & AI",
        "bullets": [
            "We support you in realizing data-driven solutions leveraging AI, to achieve results.",
            "Our Data, Analytics & AI solutions solve complex tasks tailored to your needs.",
            "Our solutions integrate into any existing environment.",
            "Our strategic technology partners are Microsoft, AWS, Databricks & Sovereign Open Source — "
            "with broad experience across many other Data & AI technologies.",
        ],
    },
    "strength": {
        "title": "A Strong Business Unit",
        "subtitle": "ELCA has a strong Business Unit on Data, Analytics & AI",
        "pillars": [
            (">30 years", ["Of experience in Data, Analytics & AI"]),
            (">75 specialists", ["Local in Switzerland: Zürich, Bern, Basel, Lausanne, Geneva"]),
            ("Nearshore & Offshore", ["Spain and Vietnam"]),
            (">100 projects", ["Delivered in Data & AI over the last three years"]),
        ],
    },
    "domains": {
        "title": "2026 Focus: Four Solution Domains",
        "subtitle": "Where ELCA DataAI concentrates to achieve results for customers",
        "pillars": [
            ("Cloud Data & Analytics", [
                "Secure cloud data platforms for banks (Finnova)",
                "IoT & real-time analytics for energy & manufacturing",
                "Pseudonymized analytics platforms for health",
                "Databricks DPaaS, operational data hubs, data migrations",
            ]),
            ("AI Applications", [
                "AI Accelerator Workshops for executives",
                "Enterprise-Grade Conversational AI",
                "Swiss Sovereign AI Platform",
                "Multi-Modal Agentic RAG · Computer Vision · Signal Processing",
            ]),
            ("Data Quality Management", [
                "Trusted, high-quality data as the foundation for every analytics and AI initiative",
            ]),
            ("Data & AI Governance", [
                "Strategy & Vision", "Governance Framework", "Data FinOps Optimization",
            ]),
        ],
    },
}

# ---------------------------------------------------------------------------
# SECTION 3 — Reference project library
# ---------------------------------------------------------------------------
# Curated from "All References OnePager" on SharePoint (PublicCorporateAssets/
# Corp Decks PPTs/2026). That source deck tracks ~129 references across the
# BU; only a handful had complete, presentable content at the time this file
# was written. ADD MORE ENTRIES HERE as champions pull them from the master
# library — that master deck states these industry totals as of writing:
# Public Administration 30, Financial Services 21, Transportation &
# Logistics 21, Retail/Energy/Manufacturing 15, Health & Life Science 15,
# Other 14, Telecom 9, Defense 4.
#
# `industries` uses lowercase-hyphen keys so it can be matched against the
# --industry argument (see build_presales_deck.py's reference-selection step).

REFERENCES = [
    {
        "client": "SBB",
        "title": "Automated Quality Validation of Survey Data",
        "industries": ["transportation", "public-sector"],
        "short_benefit": "Automated a manual survey QA step, freeing experts for complex cases with a full audit trail.",
        "benefit": (
            "Automating the final quality check of survey data — previously done by hand by "
            "Quality Managers — dramatically cut repetitive manual work and freed experts to "
            "focus on the complex cases, improving throughput and consistency for downstream "
            "analytics while keeping a full audit trail."
        ),
        "technologies": "Snowflake (Snowpark), Spring Batch, Postgres, MLflow, Argo CD/Workflow",
        "effort": "0.5 FTE since September 2025",
    },
    {
        "client": "ETH Zurich",
        "title": "BTV-Bot — AI Search Chatbot",
        "industries": ["education-research", "public-sector"],
        "short_benefit": "Self-service, source-backed answers from internal docs, cutting repetitive support requests.",
        "benefit": (
            "A RAG-based knowledge assistant that gives staff and partners self-service, "
            "source-backed answers from internal documentation — cutting repetitive support "
            "requests and making complex knowledge accessible at scale."
        ),
        "technologies": "Azure, Terraform, OpenAI, Semantic Kernel, Streamlit, Python, RAG",
        "effort": "20 person days",
    },
    {
        "client": "ETH Zurich",
        "title": "RAG-based Guidelines Chatbot",
        "industries": ["education-research", "public-sector"],
        "short_benefit": "Version-aware RAG over internal guidelines, with defensible, source-backed answers.",
        "benefit": (
            "A second RAG deployment giving self-service access to internal guidelines, with "
            "version-aware, context-sensitive retrieval and defensible, source-backed answers — "
            "reducing repetitive support load and improving onboarding."
        ),
        "technologies": "Azure (Functions, Event Grid, Storage, AI Search, Document Intelligence, "
                         "OpenAI), Terraform, Semantic Kernel, Streamlit, Python",
        "effort": "20 person days (PoC)",
    },
    {
        "client": "Swisscom",
        "title": "Teradata to AWS Migration — ODCS-based Data Governance",
        "industries": ["telecom"],
        "short_benefit": "Made ODCS the vendor-agnostic metadata backbone, so tool changes no longer fragment governance.",
        "benefit": (
            "Made the Open Data Contract Standard (ODCS) the vendor-agnostic metadata backbone "
            "for data products, ownership, quality and SLAs — so tool changes no longer fragment "
            "governance and the organisation isn't locked to one vendor's roadmap."
        ),
        "technologies": "AWS, Datahub, ODCS",
        "effort": "10 person days",
    },
    {
        "client": "Swisscom",
        "title": "ODCS and Data Catalog Integration",
        "industries": ["telecom"],
        "short_benefit": "Mapped ODCS to the data catalog and industrialised the ingestion across teams.",
        "benefit": (
            "Mapped ODCS to the Datahub data-catalog attributes and built the converter that lets "
            "ODCS-defined metadata flow straight into the catalog — industrialised across "
            "cross-functional teams."
        ),
        "technologies": "Datahub, ODCS",
        "effort": "10 person days",
    },
]

# ---------------------------------------------------------------------------
# SECTION 4 — Industry modules
# ---------------------------------------------------------------------------
# One entry per industry ELCA DataAI has a tailored slide for. The
# original brief left this open ("Industry-specific slides (?)") — the
# intended pattern is: start with the industries you pitch most often, add
# more as champions write them, and fall back gracefully (see
# build_presales_deck.py) rather than inventing content for a vertical
# nobody has actually vetted yet.
#
# Key format matches REFERENCES[i]["industries"] entries so a deck can
# eventually cross-check "do we have both an industry module AND reference
# proof points for this vertical" — right now Financial Services has a
# module but the reference library doesn't yet have a Financial Services
# entry, which is a real, honest gap the skill should surface, not paper over.

INDUSTRY_MODULES = {
    "financial-services": {
        "title": "Data, Analytics & AI for Financial Services",
        "short_title": "Financial Services",
        "subtitle": "Where banks and insurers are putting Data & AI to work",
        "pillars": [
            ("Risk & Fraud Analytics", [
                "Pattern-based fraud and anomaly detection on transaction and behavioural data",
                "Credit risk scoring with human-in-the-loop oversight",
                "Model governance to keep pace with FINMA and regulatory expectations",
            ]),
            ("Personalization & Marketing", [
                "Next-best-offer and cross-sell scoring across channels",
                "Customer affinity models feeding real-time engagement",
                "Contact-strategy tooling to avoid customer fatigue while staying compliant",
            ]),
            ("Conversational & Self-Service", [
                "Enterprise-grade, source-backed conversational AI for customers and staff",
                "Authenticated, personalized digital assistants beyond FAQ handling",
            ]),
            ("Core Data Modernization", [
                "Cloud and sovereign data platforms sized for regulated environments",
                "Augmenting legacy core banking systems incrementally, without a risky full replacement",
                "Data governance and quality as the foundation for every model above",
            ]),
        ],
    },
}
