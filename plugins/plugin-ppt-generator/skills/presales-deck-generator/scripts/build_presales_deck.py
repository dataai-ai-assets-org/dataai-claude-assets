"""
build_presales_deck.py — renders a pre-sales PowerPoint from the shared
content library plus per-call dynamic content (AI use cases, industry match,
selected references).

This script does NOT do any research or writing itself — by the time it
runs, Claude has already:
  1. researched the target company from public sources (WebSearch/WebFetch)
  2. drafted 3-4 AI use cases from that research, each with a one-line source
  3. picked an industry key (or confirmed none fits yet)
  4. picked 4-5 reference entries from content_library.REFERENCES

Those are passed in as plain Python values below. Keeping research and
writing OUT of this script is deliberate: an LLM should draft the prose,
a deterministic script should never invent it.

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
from content_library import GENERAL_ELCA, DATAAI_BL, REFERENCES, INDUSTRY_MODULES

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


def select_references(industry_key, limit=5):
    """Pick references matching the industry, falling back to the
    highest-signal cross-industry proof points when there's no exact match.
    Returns (chosen, matched_by_industry: bool) so the caller can decide
    whether to label the slide as same-industry or cross-industry proof."""
    if industry_key:
        matches = [r for r in REFERENCES if industry_key in r["industries"]]
        if matches:
            return matches[:limit], True
    # No same-industry reference — don't fabricate one. Return the
    # strongest general proof points instead and let the caller be honest
    # about it on the slide.
    return REFERENCES[:limit], False


def build_deck(company, industry_key, context, ai_use_cases, out_path,
                selected_references=None):
    """
    company: str, e.g. "PostFinance"
    industry_key: str key into INDUSTRY_MODULES / REFERENCES[i]['industries'],
                   or None if the call doesn't map to a known industry yet
    context: short free-text description of the deal (used on the title slide)
    ai_use_cases: list of dicts, each {"headline": str, "description": str, "source": str}
                  3 or 4 items — already drafted by Claude from public research
    out_path: where to save the .pptx
    selected_references: optional pre-picked list from content_library.REFERENCES;
                          if omitted, select_references() picks automatically
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
    LL(s, 18, ['About ELCA', 'Data, Analytics & AI at ELCA', industry_label,
               f'Possible AI Use Cases for {company}', 'References', 'Next Steps'])
    LL(s, 19, ['Who we are, at a glance', 'Our Business Line, capabilities and 2026 focus',
               'Where we can help', 'Ideas to discuss and validate together',
               'Proof points from recent engagements', 'How we can move forward together'])

    # 2b. Context (optional) — deal stage / what's already known about the
    # prospect, as its own slide rather than squeezed onto the title slide.
    if context:
        s = A('Text Content only 1')
        T(s, 0, f'Where We\'re Starting With {company}')
        LL(s, 1, [context])

    # 3. Chapter 01 — About ELCA
    s = A('Chapter Slide 1')
    T(s, 0, '01')
    T(s, 1, 'About ELCA')
    T(s, 14, 'An independent Swiss IT company, in business since 1968.')

    _pillars_slide(A, 'Pillars_4col', GENERAL_ELCA['glance'])
    _pillars_slide(A, 'Pillars_2col', GENERAL_ELCA['local_partner'])
    _pillars_slide(A, 'Pillars_3col', GENERAL_ELCA['group_structure'])
    _pillars_slide(A, 'Pillars_4col', GENERAL_ELCA['services'])

    # 4. Chapter 02 — Data, Analytics & AI at ELCA
    s = A('Chapter Slide 1')
    T(s, 0, '02')
    T(s, 1, 'Data, Analytics & AI')
    T(s, 14, 'We love to make you successful with Data, Analytics & AI.')

    s = A('Text Content only 1')
    T(s, 0, DATAAI_BL['mission']['title'])
    LL(s, 1, DATAAI_BL['mission']['bullets'])

    _pillars_slide(A, 'Pillars_4col', DATAAI_BL['strength'])
    _pillars_slide(A, 'Pillars_4col', DATAAI_BL['domains'])

    # 5. Chapter 03 — Industry module (or an honest gap slide)
    # Chapter title must stay short (one line) — the layout's tagline (idx14)
    # sits at a fixed position right below it, same risk as the title slide.
    s = A('Chapter Slide 1')
    T(s, 0, '03')
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

    # 6. AI Use Cases — the one genuinely dynamic content slide
    s = A('Pillars_4col' if len(ai_use_cases) >= 4 else 'Pillars_3col')
    T(s, 0, f'Possible AI Use Cases for {company}')
    T(s, 1, 'Draft ideas for discussion — grounded in public information, to validate together')
    idx_pairs = [(10, 11), (12, 13), (14, 15), (16, 17)][: len(ai_use_cases)]
    for (title_idx, body_idx), uc in zip(idx_pairs, ai_use_cases):
        T(s, title_idx, uc['headline'])
        LL(s, body_idx, [uc['description'], '', f"Source: {uc['source']}"])

    # 7. References — kept to short clauses (short_benefit, not the full
    # `benefit` text) and capped at 4 per slide. A flowing bullet list with
    # full-sentence benefits for 4-5 references reliably overflows this
    # layout; resist the temptation to add more detail here than that.
    if selected_references is None:
        selected_references, matched = select_references(industry_key, limit=4)
    else:
        matched = True
    match_label = 'same-industry' if matched else 'cross-industry — no same-industry reference yet'
    s = A('Text Content only 1')
    T(s, 0, f'References ({match_label})')
    LL(s, 1, [f"{r['client']} — {r['title']}: {r.get('short_benefit', r['benefit'])}"
              for r in selected_references])

    # 8. Contact
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
