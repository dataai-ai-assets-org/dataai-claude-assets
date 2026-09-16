"""
Content library for the presales-deck-generator skill.

This skill is deliberately thin on maintained content. Its job is to take
whatever Claude researches about the prospect for a given call — company,
industry, AI use cases — and render it in ELCA style. General ELCA facts and
Data & AI Business Line content are NOT part of this skill; nothing about
"who ELCA is" is needed to produce this deck.

The one thing worth maintaining centrally is the industry-module framework
below: ELCA's own curated pitch angle per vertical, not something a web
search would reliably reconstruct the same way twice. build_presales_deck.py
imports this file directly.

Reference-project slides are also NOT maintained here. They're merged in per
call from a reference deck the user points to in their prompt — see
"Dependency: deck-merger" and Step 5 in presales-deck-generator/SKILL.md.

WHO EDITS THIS FILE: whoever ELCA's Data & AI BL nominates as skill
champions (see the department's shared plugin marketplace). No engineering
background is required — it's plain Python dicts, organised by clearly
labelled section. Keep entries short: this content is meant to be read off
a slide, not a report.
"""

# ---------------------------------------------------------------------------
# Industry modules
# ---------------------------------------------------------------------------
# One entry per industry ELCA DataAI has a tailored slide for. The
# original brief left this open ("Industry-specific slides (?)") — the
# intended pattern is: start with the industries you pitch most often, add
# more as champions write them, and fall back gracefully (see
# build_presales_deck.py) rather than inventing content for a vertical
# nobody has actually vetted yet.
#
# Keys are lowercase-hyphenated (e.g. "financial-services") so they can also
# be matched against client/text hints when picking reference slides from the
# user-supplied reference deck — see Step 5 in presales-deck-generator/SKILL.md.

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
