"""
build_presales_deck.py — renders a pre-sales PowerPoint from the shared
content library plus per-call dynamic content (AI use cases, industry match).

This script does NOT do any research or writing itself — by the time it
runs, Claude has already:
  1. researched the target company from public sources (WebSearch/WebFetch)
  2. drafted 3-4 AI use cases from that research, each with a one-line source
  3. picked an industry key (or confirmed none fits yet)

Those are passed in as plain Python values below. Keeping research and
writing OUT of this script is deliberate: an LLM should draft the prose,
a deterministic script should never invent it.

This script does NOT produce reference-project slides either. It only adds a
"References" chapter-divider slide as an anchor point. The actual reference
slides are spliced in afterwards, verbatim, from a reference deck the user
supplies — see "Dependency: deck-merger" and Step 5 in
presales-deck-generator/SKILL.md.

Usage: edit the CONFIG block at the bottom (or import build_deck() from
another script) and run with python3.
"""

import sys
import os

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)

# Path to the elca-pptx skill's helpers — this skill depends on elca-pptx
# for correct ELCA-branded rendering (see SKILL.md, "Dependency: elca-pptx").
#
# Both skills ship together in the plugin-ppt-generator plugin, so the
# common case needs no configuration: elca-pptx is a sibling directory under
# the plugin's skills/ folder and is found relative to this file. The other
# candidates keep the script working when the skills are installed
# standalone. Resolution stops at the first directory that actually contains
# the helpers, so a stale env var or leftover legacy path can't silently win
# over a real installation.
def _find_elca_pptx_skill_dir():
    plugin_root = os.environ.get("CLAUDE_PLUGIN_ROOT")
    candidates = [
        os.environ.get("ELCA_PPTX_SKILL_DIR"),
        os.path.join(plugin_root, "skills", "elca-pptx") if plugin_root else None,
        os.path.abspath(os.path.join(SCRIPT_DIR, "..", "..", "elca-pptx")),
        "/root/.claude/skills/synced/elca-pptx",
    ]
    checked = []
    for candidate in candidates:
        if not candidate:
            continue
        checked.append(candidate)
        if os.path.isfile(os.path.join(candidate, "scripts", "pptx_helpers.py")):
            return candidate
    raise RuntimeError(
        "Could not locate the elca-pptx skill (looked for scripts/pptx_helpers.py in: "
        + ", ".join(checked)
        + "). This skill renders through elca-pptx's template and helpers and "
        "should not fall back to hand-rolled python-pptx calls. Set "
        "ELCA_PPTX_SKILL_DIR to the elca-pptx skill directory."
    )


ELCA_PPTX_SKILL_DIR = _find_elca_pptx_skill_dir()
sys.path.insert(0, os.path.join(ELCA_PPTX_SKILL_DIR, "scripts"))

from pptx import Presentation
from pptx_helpers import T, LL
from content_library import INDUSTRY_MODULES

TEMPLATE = os.path.join(ELCA_PPTX_SKILL_DIR, "ELCA PPT Template.pptx")


def _pillars_slide(A, layout_name, block, extra_pillars=None):
    """Render a dict shaped like content_library's pillar blocks onto a
    Pillars_Ncol layout. `layout_name` must match the pillar count."""
    s = A(layout_name)
    T(s, 0, block["title"])
    T(s, 1, block["subtitle"])
    pillars = block["pillars"]
    idx_pairs = [(10, 11), (12, 13), (14, 15), (16, 17), (18, 19)][: len(pillars)]
    for (title_idx, body_idx), (p_title, p_body) in zip(idx_pairs, pillars):
        T(s, title_idx, p_title)
        LL(s, body_idx, p_body)
    return s


def build_deck(company, industry_key, context, ai_use_cases, out_path):
    """
    company: str, e.g. "PostFinance"
    industry_key: str key into INDUSTRY_MODULES, or None if the call doesn't
                   map to a known industry yet
    context: short free-text description of the deal (used on the title slide)
    ai_use_cases: list of dicts, each {"headline": str, "description": str, "source": str}
                  3 or 4 items — already drafted by Claude from public research
    out_path: where to save the .pptx

    Does not take reference-project content — see the module docstring and
    Step 5 in presales-deck-generator/SKILL.md for how those get merged in
    afterwards from a user-supplied reference deck.
    """
    prs = Presentation(TEMPLATE)
    ORIG = len(prs.slides)
    layouts = {lay.name: lay for lay in prs.slide_master.slide_layouts}

    def A(name):
        return prs.slides.add_slide(layouts[name])

    # 1. Title — keep the title itself short and fixed (it's the one line that
    # must never wrap, since idx1 sits directly below it), and keep idx1 to
    # just the company name. idx1 is sized for a short "two-line heading" per
    # the template's own placeholder text, and idx14 (the smaller placeholder
    # right below it) has almost no gap above it — even a short context
    # sentence there visibly collides with idx1's text (verified the hard
    # way). The context sentence gets its own slide instead (see 2b below),
    # where it has room to read as a real sentence.
    s = A('Title Slide')
    T(s, 0, 'Data, Analytics & AI')
    T(s, 1, company)

    # 2. Agenda
    s = A('Agenda 1')
    T(s, 0, 'Agenda')
    industry_label = INDUSTRY_MODULES.get(industry_key, {}).get(
        'title', f'{company} — Context & Opportunities'
    )
    LL(s, 18, [industry_label, f'Possible AI Use Cases for {company}',
               'References', 'Next Steps'])
    LL(s, 19, ['Where we can help', 'Ideas to discuss and validate together',
               'Proof points from recent engagements', 'How we can move forward together'])

    # 2b. Context (optional) — deal stage / what's already known about the
    # prospect, as its own slide rather than squeezed onto the title slide.
    if context:
        s = A('Text Content only 1')
        T(s, 0, f'Where We\'re Starting With {company}')
        LL(s, 1, [context])

    # 3. Chapter 01 — Industry module (or an honest gap slide). This is the
    # deck's first content chapter: the skill carries no general ELCA or
    # Data & AI Business Line content, only what's specific to this call.
    # Chapter title must stay short (one line) — the layout's tagline (idx14)
    # sits at a fixed position right below it, same risk as the title slide.
    s = A('Chapter Slide 1')
    T(s, 0, '01')
    module = INDUSTRY_MODULES.get(industry_key)
    if module:
        short_title = module.get('short_title') or (industry_key or company).replace('-', ' ').title()
        T(s, 1, short_title)
        T(s, 14, module['subtitle'])
        _pillars_slide(A, 'Pillars_4col', module)
    else:
        T(s, 1, company)
        T(s, 14, 'No tailored industry module yet for this vertical — filled in for this call.')
        s = A('Text Content only 1')
        T(s, 0, f'What We Understand About {company}')
        LL(s, 1, [context] if context else ['Add what you know about the prospect here.'])

    # 4. AI Use Cases — the one genuinely dynamic content slide
    s = A('Pillars_4col' if len(ai_use_cases) >= 4 else 'Pillars_3col')
    T(s, 0, f'Possible AI Use Cases for {company}')
    T(s, 1, 'Draft ideas for discussion — grounded in public information, to validate together')
    idx_pairs = [(10, 11), (12, 13), (14, 15), (16, 17)][: len(ai_use_cases)]
    for (title_idx, body_idx), uc in zip(idx_pairs, ai_use_cases):
        T(s, title_idx, uc['headline'])
        LL(s, body_idx, [uc['description'], '', f"Source: {uc['source']}"])

    # 5. References — a chapter divider only. This script never renders
    # reference content itself; the actual slides get spliced in right after
    # this divider, verbatim, from a reference deck the user supplies (see
    # Step 5 in SKILL.md). Run inspect_deck.py on the saved output to find
    # this slide's exact position before calling merge_decks.py — it shifts
    # depending on whether the context and industry-module slides above ran.
    s = A('Chapter Slide 1')
    T(s, 0, '02')
    T(s, 1, 'References')
    T(s, 14, 'Proof points from recent engagements.')

    # 6. Contact
    s = A('Final/Contact Slide')
    T(s, 0, "Let's Talk.")
    T(s, 1, f'Ready to explore what Data, Analytics & AI can do for {company}?')
    T(s, 14, 'Markus Grob')
    T(s, 16, 'Head of Business Line: Data, Analytics & AI, ELCA')
    T(s, 17, 'markus.grob@elca.ch')
    T(s, 18, 'www.elca.ch/our-expertise/data-analytics-ai')

    # Rebuild sldIdLst — required whenever building fresh from the template,
    # see elca-pptx SKILL.md Step 5.
    all_ids = list(prs.slides._sldIdLst)
    our_ids = all_ids[ORIG:]
    sldIdLst = prs.slides._sldIdLst
    for sid in all_ids:
        sldIdLst.remove(sid)
    for sid in our_ids:
        sldIdLst.append(sid)

    prs.save(out_path)
    return out_path


if __name__ == '__main__':
    # Example / smoke-test invocation — replace with a real call, or import
    # build_deck() from a per-run script after research is done.
    demo_use_cases = [
        {
            "headline": "Example Use Case",
            "description": "Replace this with a real, researched use case before running for a real prospect.",
            "source": "n/a",
        },
    ]
    build_deck(
        company="Example Corp",
        industry_key=None,
        context="Pre-Sales Introduction",
        ai_use_cases=demo_use_cases,
        out_path=os.path.join(SCRIPT_DIR, "..", "demo_output.pptx"),
    )
    print("Wrote demo_output.pptx")
