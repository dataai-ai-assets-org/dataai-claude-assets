---
name: presales-deck-generator
description: "Generates a first-pre-sales-call PowerPoint for ELCA's Data, Analytics & AI Business Line, given a company name plus optional industry/context. Pulls together general ELCA slides, Data & AI BU slides, an industry module when one exists, a company-specific AI-use-case slide drafted from public research, and 4-5 reference project slides. Use this skill whenever someone asks to build, generate, or put together a pre-sales deck, sales presentation, or pitch deck for a named prospect or customer in the Data & AI space — even if they just say something like 'can you make me a deck for our call with Acme AG next week' or 'I need slides for the first meeting with Acme'. Also use it when someone wants to update or regenerate an existing pre-sales deck for a different company, or asks what a pre-sales deck for a given industry would look like."
---

# Pre-Sales Deck Generator

Builds the deck ELCA's Data, Analytics & AI Business Line uses to open a first
conversation with a prospect: who ELCA is, what the BU does, what's specific
to the prospect's industry, a few AI ideas worth discussing tailored to that
one company, and proof points from real engagements.

## Why this skill is shaped the way it is

Most of a pre-sales deck should say the same thing every time — ELCA's
facts don't change between calls, and the BU's capabilities don't either.
Regenerating that content per call wastes time and risks drifting off
message. So this skill treats the deck as **mostly a maintained content
library, with one genuinely dynamic slide**: the AI use cases, which have to
be specific to the company in front of you or they're not worth including.

Concretely: `scripts/content_library.py` holds everything that's the same
every time (general ELCA facts, the BU's mission/stats/domains, industry
modules, reference projects). `scripts/build_presales_deck.py` assembles a
deck from that library plus the handful of things that change per call. Your
job when this skill triggers is to gather those per-call things, then run
the script — not to write slide content from scratch each time.

## Dependency: elca-pptx

This skill renders through the **elca-pptx** skill's template and helpers
(`T()`/`LL()`) — it does not draw slides itself. Both skills ship together
in the **plugin-ppt-generator** plugin precisely so this dependency is
always satisfied: elca-pptx sits next to this skill under the plugin's
`skills/` directory, and `build_presales_deck.py` finds it there on its own.

Read elca-pptx's SKILL.md before your first build anyway. The rendering
rules there (never `clear()`/`add_run()`/`font.*` on layout placeholders)
are what keep pillar titles red and bullets intact, and you'll need them the
moment a deck needs a slide the build script doesn't already produce.

If the helpers genuinely can't be found, `build_presales_deck.py` raises an
error listing where it looked. Fix the path (set `ELCA_PPTX_SKILL_DIR`)
rather than hand-rolling python-pptx calls — that's how decks end up with
black pillar titles and broken bullets.

## Workflow

### Step 1 — Gather the per-call inputs

You need, at minimum, a **company name**. Ask for it if it's missing —
there's no reasonable default. Also try to get:

- **Industry** (e.g. "financial services", "public sector", "retail"). If
  the user doesn't give one, infer it from context or a quick search on the
  company. Match it to a key in `INDUSTRY_MODULES` in
  `scripts/content_library.py` — the current keys are listed at the top of
  that dict. Getting this right matters: it decides which industry module
  and which reference projects the deck uses.
- **Context**: deal stage, what's already known about the prospect's
  situation, who's running the call. This becomes the title-slide subtitle
  and, when there's no industry module, the fallback "what we understand"
  slide.

If the user is clearly in a hurry ("just generate something now"), proceed
with your best inference and say what you assumed, rather than blocking on
questions nobody's there to answer.

### Step 2 — Research the company publicly

Before drafting anything, search the web for what's publicly known about the
company: what they do, their scale/segment, anything recently public about
their strategy, priorities, or pain points. Keep a running list of facts
with their sources (URL + a one-line description of what it says) — you'll
need these to ground the use-case slide honestly.

This step is not optional. An AI use-case slide with no real grounding is
worse than no use-case slide — it reads as generic and erodes trust in the
first five minutes of a sales call.

### Step 3 — Draft 3-4 AI use cases

From the research, draft 3-4 plausible AI use cases specific to this
company, in ELCA's voice (data-driven, results-oriented, partnership
framing — see `DATAAI_BU['mission']` in the content library for the tone).
Each use case is a dict:

```python
{"headline": "Short, punchy title (≤ ~25 characters)",
 "description": "ONE sentence, ~20-25 words: what it is, why it matters to them.",
 "source": "What public fact this is grounded in, in ~8-10 words — e.g. "
           "'PostFinance's own AI page, on legacy augmentation'"}
```

Keep `description` and `source` genuinely short — this renders as one of four
columns on a single slide alongside three other use cases, and there is no
autofit safety net: a two-sentence description with a full-sentence source
reliably overflows past the bottom of the slide (verified the hard way while
building this skill). If you find yourself writing two sentences, cut it to
one and move any nuance into the delivery conversation, not the slide.

`headline` is character-count-limited, not word-count-limited — this column
is narrow, and a headline that wraps to 3 lines collides with the description
directly below it (that placeholder has no autofit either). "Harden Fraud &
Risk Models" (27 characters) wraps cleanly to 2 lines; "Modernize Without
Ripping Out Core Systems" (44 characters) wraps to 3 and visibly overlaps the
bullet below it — verified the hard way in this skill's own demo run. Stay at
or under ~25-28 characters, and count characters rather than words, especially
with compound terms like "Next-Best-Offer" that can't break mid-word.

Rules for this step, because it's the one place this skill can go wrong in
a way that damages trust with a real customer:

- **If you can't find enough public information to ground a use case,
  don't invent one to fill the slide.** Fewer, well-grounded use cases beat
  four generic ones. Three is a fine number; two is acceptable; don't pad.
- Every use case needs a `source` a human could actually check. "Industry
  best practice" is not a source — a specific page, statement, or fact is.
- Tie each use case back to one of the BU's actual solution domains
  (`DATAAI_BU['domains']`) where you can — it should read as something ELCA
  is positioned to deliver, not a generic AI idea anyone could pitch.

### Step 4 — Pick reference projects

Call `select_references(industry_key)` from `build_presales_deck.py` — it
matches references to the industry and tells you (`matched_by_industry`)
whether it found a same-industry proof point or had to fall back to
cross-industry ones. **Don't override this to hide a gap.** If there's no
same-industry reference yet, the deck should say "cross-industry proof
points" rather than imply a same-industry relationship that doesn't exist —
`build_presales_deck.py` handles that labeling automatically as long as you
don't hand it a hand-picked `selected_references` list that misrepresents
the match.

### Step 5 — Build the deck

```python
import sys, os
# Inside the plugin-ppt-generator plugin, ${CLAUDE_PLUGIN_ROOT} points at
# the plugin directory; standalone, use this skill's own directory.
sys.path.insert(0, os.path.join(
    os.environ.get("CLAUDE_PLUGIN_ROOT", "<path-to-this-skill>/.."),
    "skills", "presales-deck-generator", "scripts"))
from build_presales_deck import build_deck

build_deck(
    company="Acme AG",
    industry_key="financial-services",   # or None — see content_library.py for valid keys
    context="First call, referred in by their CTO; interested in fraud analytics.",
    ai_use_cases=[...],                   # from Step 3
    out_path="/path/to/output/Acme_AG_Presales.pptx",
)
```

### Step 6 — Visually QA every slide

Same discipline as elca-pptx itself — don't skip this because the content
came from a script. The point is to look at rendered slides, not to trust
that the script placed text sensibly; overflow is invisible in the XML.

Render the deck to images however the machine allows. In a Linux
environment with the public `pptx` skill and poppler available:

```bash
python3 /path/to/skills/pptx/scripts/office/soffice.py --headless --convert-to pdf output.pptx
pdftoppm -jpeg -r 110 output.pdf slide
```

On a local Windows install, neither of those is usually present. Use
LibreOffice directly if it's installed:

```powershell
& "C:\Program Files\LibreOffice\program\soffice.exe" --headless --convert-to pdf --outdir . output.pptx
```

and convert the PDF pages to images with whatever is available (`pdftoppm`
from a poppler build, ImageMagick, or `pypdfium2` from Python). If nothing
can render, say so and ask the person to open the deck and eyeball it — an
unreviewed deck going to an account team is worse than an admitted gap.

Then `Read` each slide image and check: no text overflow (long company
names or industry module titles can push pillar titles to 3 lines — shorten
if so, the same way you'd fix any elca-pptx overflow), pillar titles are
red, the AI use-case slide's source lines are actually present and legible,
and the References slide correctly says same-industry vs. cross-industry.

### Step 7 — Deliver as a draft, always

This deck is a **first draft for the account team to review**, never
something to send straight to a customer. Say so explicitly when you hand
it over — the AI use-case slide in particular needs a human sanity check
before it's in front of the prospect, however well-sourced.

## Maintaining the content library

`scripts/content_library.py` is the single source of truth for everything
that isn't per-call. It's plain, commented Python — no engineering
background needed to edit it. When ELCA's general facts, the BU's stats, or
an industry module need updating, edit that file directly rather than
hard-coding a one-off change in a build script; every future deck should
benefit from the update, not just the one you're building today.

This skill ships in the **plugin-ppt-generator** plugin alongside
elca-pptx, so BU champions — not just whoever happens to be running this
session — can maintain `content_library.py` over time. Edits go to the
plugin repository and reach everyone on the next `/plugin update`. See the
plugin's commit history for what's changed and when re-pulling from the
source SharePoint decks (2026 Corporate Deck, DataAI BU intro slides, All
References OnePager — all under PublicCorporateAssets/Corp Decks PPTs/2026)
is overdue.

Adding a new industry module: add an entry to `INDUSTRY_MODULES` following
the existing `financial-services` example — a title, subtitle, and up to
four pillars of (heading, bullet-list) tuples. Adding a reference: add an
entry to `REFERENCES` with real content pulled from a real engagement —
never fabricate a reference project.
