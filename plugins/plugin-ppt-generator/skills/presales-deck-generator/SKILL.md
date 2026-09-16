---
name: presales-deck-generator
description: "Generates a first-pre-sales-call PowerPoint for ELCA's Data, Analytics & AI Business Line, given a company name plus optional industry/context. Researches the prospect live and pulls together an industry module when one exists, a company-specific AI-use-case slide drafted from public research, and reference-project slides merged in from a reference deck the user points to — no general ELCA or Business Line content, just what's specific to this call, rendered in ELCA style. Use this skill whenever someone asks to build, generate, or put together a pre-sales deck, sales presentation, or pitch deck for a named prospect or customer in the Data & AI space — even if they just say something like 'can you make me a deck for our call with Acme AG next week' or 'I need slides for the first meeting with Acme'. Also use it when someone wants to update or regenerate an existing pre-sales deck for a different company, or asks what a pre-sales deck for a given industry would look like."
---

# Pre-Sales Deck Generator

Builds the deck ELCA's Data, Analytics & AI Business Line uses to open a first
conversation with a prospect: what's specific to the prospect's industry, a
few AI ideas worth discussing tailored to that one company, and proof points
from real engagements — rendered in ELCA style.

## Why this skill is shaped the way it is

This skill is deliberately not a "corporate deck generator." It carries no
general ELCA content and no Data & AI Business Line content — no "who we
are" chapter, no capability/mission slides. Every real pre-sales
conversation needs those covered elsewhere (a corporate deck, a leave-behind,
whatever the account team already uses); baking a static copy into this
skill would just be one more place for that messaging to drift out of sync
with the real thing.

What this skill actually does is narrow and live: take a company name (and
optional industry/context), research the prospect on the public web, draft
AI ideas grounded in that research, and render the result in ELCA's visual
style via elca-pptx. The only thing worth maintaining centrally is the
industry-module framework in `scripts/content_library.py` — ELCA's own
curated pitch angle per vertical, which a web search wouldn't reliably
reconstruct the same way twice. Reference-project slides get the same
treatment as everything else here: nothing pre-loaded, every call gets its
reference slides fresh, merged verbatim from a deck the user points to — see
Step 5.

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

## Dependency: deck-merger

Reference-project slides come from the **deck-merger** skill, which ships in
the separate **plugin-deck-merger** plugin — not bundled here the way
elca-pptx is. Check it's installed (`/plugin list`) before Step 5; if it
isn't, tell the user and skip references for this call rather than
fabricating slide content by hand.

deck-merger's `inspect_deck.py` and `merge_decks.py` are what Step 5 uses.
Read deck-merger's own SKILL.md if you haven't used it before — the
principles that matter here are the same ones it's built around: copy slide
XML verbatim rather than rebuild it, and never silently hide a gap (a hidden
slide, a housekeeping slide, no matching industry) from the user.

## Workflow

### Step 1 — Gather the per-call inputs

You need, at minimum, a **company name**. Ask for it if it's missing —
there's no reasonable default. Also try to get:

- **Industry** (e.g. "financial services", "public sector", "retail"). If
  the user doesn't give one, infer it from context or a quick search on the
  company. Match it to a key in `INDUSTRY_MODULES` in
  `scripts/content_library.py` — the current keys are listed at the top of
  that dict. Getting this right matters: it decides which industry module
  the deck uses, and guides which reference slides you pick in Step 5.
- **A reference deck**, if the user gave a path or link to one in their
  prompt. If they didn't, ask once; if they say they don't have one handy,
  proceed without references for this call (Step 5 covers what to do then).
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
company, in ELCA's voice: data-driven, results-oriented, partnership
framing — a peer offering to help solve a real problem, not a vendor
pitching a generic capability. Each use case is a dict:

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
- Each use case should read as something ELCA's Data & AI Business Line is
  actually positioned to deliver (cloud data platforms, AI applications,
  data quality, data governance) — not a generic AI idea anyone could pitch.

### Step 4 — Build the base deck

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
    out_path="/path/to/output/Acme_AG_Presales_base.pptx",
)
```

This deck ends with a "References" chapter divider (title "02") followed
immediately by the Contact slide — no reference content yet. That's Step 5.

### Step 5 — Merge in reference-project slides

This step needs the **deck-merger** skill (plugin-deck-merger) and the
reference deck path/link gathered in Step 1. If there's no reference deck for
this call, skip straight to Step 6 and say plainly in your response that the
deck has no reference slides — never invent one to avoid an empty section.

1. **Inventory the reference deck.** Run deck-merger's
   `inspect_deck.py` on it to see every slide's position, layout name, and
   opening text (hidden and housekeeping slides included — flag those, same
   as deck-merger always does).
2. **Pick slides matching the industry.** Read the inventory for slides
   whose client, title, or text preview matches `industry_key` (or the
   company's sector generally). Prefer 3-5 strong, specific proof points
   over padding with weak ones. If nothing matches the industry, pick the
   strongest general proof points instead and say so explicitly when you
   hand the deck over — "cross-industry proof points, no same-industry
   reference yet" — rather than presenting them as same-industry.
3. **Locate the insertion point.** Run deck-merger's `inspect_deck.py` on
   the *base deck from Step 4* to find the exact 1-based position of the
   "References" chapter-divider slide (title "02") and the slide right
   after it (the Contact slide). These positions shift depending on whether
   the context slide (Step 1) and a real industry module (vs. the fallback
   gap slide) were included, so always look them up on the actual file —
   never assume fixed numbers.
4. **Splice the picked slides in with `merge_decks.py`**, using the base
   deck twice to keep everything before and after the insertion point:
   ```bash
   python3 <deck-merger-skill>/scripts/merge_decks.py -o final.pptx \
       --deck base.pptx       --slides 1-<divider_position> \
       --deck reference.pptx  --slides <picked-ranges> \
       --deck base.pptx       --slides <contact_position>
   ```
   `base.pptx` stays the first `--deck` overall so its master/theme/size
   govern the output — the imported reference slides keep inheriting their
   own master via the same mechanism deck-merger always uses.
5. Validate the result the way deck-merger's own SKILL.md describes before
   handing it over.

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
the merged reference slides render correctly (fonts/theme intact, nothing
cut off), and your response to the user correctly says same-industry vs.
cross-industry vs. no references for this call.

### Step 7 — Deliver as a draft, always

This deck is a **first draft for the account team to review**, never
something to send straight to a customer. Say so explicitly when you hand
it over — the AI use-case slide in particular needs a human sanity check
before it's in front of the prospect, however well-sourced.

## Maintaining the content library

`scripts/content_library.py` only holds `INDUSTRY_MODULES` — there's no
general ELCA or Business Line content in this skill to maintain. It's plain,
commented Python — no engineering background needed to edit it. When an
industry module needs adding or updating, edit that file directly rather
than hard-coding a one-off change in a build script; every future deck for
that industry should benefit from the update, not just the one you're
building today.

This skill ships in the **plugin-ppt-generator** plugin alongside
elca-pptx, so BL champions — not just whoever happens to be running this
session — can maintain `content_library.py` over time. Edits go to the
plugin repository and reach everyone on the next `/plugin update`.

Adding a new industry module: add an entry to `INDUSTRY_MODULES` following
the existing `financial-services` example — a title, subtitle, and up to
four pillars of (heading, bullet-list) tuples.

There is nothing to maintain here for references — that content isn't in
this file. The nearest thing to a reference library is whatever deck the
user points to per call (ELCA's "All References OnePager" on SharePoint,
under the same PublicCorporateAssets/Corp Decks PPTs/2026 folder, is the
usual one). If that link changes or a better one becomes the standard,
that's a fact worth the user telling Claude at the start of a session, not
a change to this file.
