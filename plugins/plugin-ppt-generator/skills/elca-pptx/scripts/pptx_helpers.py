"""
ELCA PPTX Helpers
=================
Layout-aware text helpers for building ELCA-branded PowerPoint slides.

Critical rule: NEVER use text_frame.clear(), p.add_run(), or run.font.*
These create an explicit <a:rPr> that overrides inherited layout styling
(colors, fonts, bullets all break). Always use T() and LL() below.

Usage:
    from pptx import Presentation
    from scripts.pptx_helpers import ph, T, LL, A

    prs = Presentation('ELCAi PPT Template.pptx')
    layouts = {lay.name: lay for lay in prs.slide_master.slide_layouts}
    def A(name): return prs.slides.add_slide(layouts[name])

    s = A('Pillars_2col')
    T(s, 0, 'My Slide Title')          # title placeholder
    T(s, 1, 'Subtitle tagline')        # subtitle — italic, smaller, auto
    T(s, 10, 'Left Pillar Header')     # red, bold, no bullet — auto from layout
    LL(s, 11, ['Bullet 1', 'Bullet 2', '', 'After spacer'])  # bullets — auto
    T(s, 12, 'Right Pillar Header')
    LL(s, 13, ['Item A', 'Item B'])

    prs.save('output.pptx')
"""
from pptx.oxml.ns import qn
from lxml import etree
import copy


def ph(slide, idx):
    """Get placeholder by placeholder_format.idx."""
    for p in slide.placeholders:
        if p.placeholder_format.idx == idx:
            return p
    return None


def _get_layout_para_template(slide, idx):
    """
    Returns (pPr_copy, rPr_copy) extracted from the layout placeholder at idx.

    This is the heart of the approach: the layout defines ALL visual styling
    (color, font, size, bullets). By copying pPr (paragraph properties, incl.
    buNone for title placeholders) and rPr (run properties: color, font, bold,
    italic, size) from the layout, we guarantee the slide inherits exactly what
    the designer intended.

    Falls back to endParaRPr when no runs exist — this happens for subtitle-type
    placeholders which store their styling (italic, Calibri, muted color) only
    in endParaRPr rather than in an actual run.
    """
    layout = slide.slide_layout
    for lph in layout.placeholders:
        if lph.placeholder_format.idx == idx:
            paras = lph.text_frame._txBody.findall(qn('a:p'))
            if not paras:
                return (None, None)
            lp = paras[0]
            pPr = lp.find(qn('a:pPr'))
            runs = lp.findall(qn('a:r'))
            rPr = runs[0].find(qn('a:rPr')) if runs else None
            # Fallback: subtitle and some placeholders store style in endParaRPr only
            if rPr is None:
                endParaRPr = lp.find(qn('a:endParaRPr'))
                if endParaRPr is not None:
                    rPr = etree.Element(qn('a:rPr'))
                    for attr, val in endParaRPr.attrib.items():
                        rPr.set(attr, val)
                    for child in endParaRPr:
                        rPr.append(copy.deepcopy(child))
            return (
                copy.deepcopy(pPr) if pPr is not None else None,
                copy.deepcopy(rPr) if rPr is not None else None,
            )
    return (None, None)


def T(slide, idx, text):
    """
    Set a single line of text on a placeholder.

    Copies pPr + rPr from the layout so colors, fonts, bullet settings,
    and bold/italic are all preserved exactly as the designer intended.
    Safe for title, subtitle, pillar header, chapter number, and tagline placeholders.
    """
    p = ph(slide, idx)
    if not p:
        return
    pPr_t, rPr_t = _get_layout_para_template(slide, idx)
    txBody = p.text_frame._txBody
    for par in txBody.findall(qn('a:p')):
        txBody.remove(par)
    new_p = etree.SubElement(txBody, qn('a:p'))
    if pPr_t is not None:
        new_p.append(pPr_t)
    r_el = etree.SubElement(new_p, qn('a:r'))
    if rPr_t is not None:
        r_el.append(rPr_t)
    etree.SubElement(r_el, qn('a:t')).text = text


def LL(slide, idx, lines):
    """
    Set multi-line text on a placeholder.

    Each string in `lines` becomes a separate paragraph inheriting the layout's
    pPr (bullet char, indent, spacing) and rPr (font, color, size). Use an
    empty string '' to insert a visual spacer paragraph between sections.

    Safe for pillar body, content body, agenda item, and any bulleted area.
    """
    p = ph(slide, idx)
    if not p:
        return
    pPr_t, rPr_t = _get_layout_para_template(slide, idx)
    txBody = p.text_frame._txBody
    for par in txBody.findall(qn('a:p')):
        txBody.remove(par)
    for text in lines:
        new_p = etree.SubElement(txBody, qn('a:p'))
        if pPr_t is not None:
            new_p.append(copy.deepcopy(pPr_t))
        r_el = etree.SubElement(new_p, qn('a:r'))
        if rPr_t is not None:
            r_el.append(copy.deepcopy(rPr_t))
        etree.SubElement(r_el, qn('a:t')).text = text
