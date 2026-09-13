---
name: elca-pptx
description: >
  Build a native ELCA-branded PowerPoint presentation using the ELCA or ELCAi
  PPT Template. Use this skill whenever the user asks to create, rebuild, or
  add slides to a PPTX — for Sounding Boards, client decks, training slides,
  strategy presentations, or any deck that should look like ELCA material.
  Also use it whenever a previous attempt produced wrong colors, missing
  bullets, black text where red was expected, or overridden fonts. The skill
  provides the complete step-by-step workflow and the Python helpers (T, LL,
  SHAPETB, BLANKTITLE) that are the safe way to set text without overriding
  layout styling.
---

# ELCA PPTX Skill

Build properly styled ELCA PowerPoint presentations — correct fonts, colors, bullets — in one pass, without fighting the template.

---

## The one rule that matters — and where it applies

**Never use `text_frame.clear()`, `p.add_run()`, or `run.font.*` on layout placeholder text frames.**

Layout placeholders (the things you access via `slide.placeholders[n]`) have all their visual design — color, font, size, bullets, indents — inherited from the layout XML. Using `add_run()` or `font.*` creates an explicit `<a:rPr>` that overrides that inheritance. Result: pillar titles lose their red color, subtitles lose italics, bullets appear where there should be none.

The `T()` and `LL()` functions in `scripts/pptx_helpers.py` are the only safe way to set text on layout placeholders. Use them exclusively for any slide that uses a named layout (Pillars, Chapter, Text Content, etc.).

**This rule does NOT apply to custom shapes.** When you create a rectangle, text box, or any shape with `add_shape()` or `add_textbox()`, it has no inherited layout styling. Setting `run.font.name`, `run.font.size`, `run.font.color.rgb` etc. is perfectly correct and required. Use `add_run()` freely on custom shapes.

---

## Two templates — choose the right one

Both templates are bundled in this skill's directory alongside `SKILL.md` and `scripts/`.

### Which template to use

| Deck is about... | Template |
|---|---|
| Agentic Engineering unit, ELCAi Method, BMAD, Claude Code, agentic PoC/offer, RhB/PostFinance agentic proof points, unit training, Sounding Board | **`ELCAi PPT Template.pptx`** — ELCAi logo bottom right |
| Everything else: general ELCA offers, client presentations, strategy, corporate, non-AE topics | **`ELCA PPT Template.pptx`** — ELCA logo bottom right |

When in doubt, default to **ELCA**. Use ELCAi only when the content is explicitly about the Agentic Engineering unit or its methodology.

### Template details

| | ELCAi | ELCA |
|---|---|---|
| File | `ELCAi PPT Template.pptx` | `ELCA PPT Template.pptx` |
| Logo | ELCAi (bottom right) | ELCA (bottom right) |
| Layouts | 32 | 43 |
| Pre-existing slides | 4 | 2 |

The ELCA template has all the ELCAi layouts plus 12 extras: `Chapter Slide Dark lines` (4 variants), `1_Final/Contact Slide`, `Contact Slide without picture`, `Agenda 2`, `Text on the right 2+3`, `Text on the left 2+3`, `Blank Slide 2`.

### How to find the bash path

When the skill loads, the system prompt's "Shell access" section lists the exact bash path for the skills directory, for example:
```
C:\Users\<username>\AppData\Roaming\Claude\...\skills → /sessions/<session-name>/mnt/.claude/skills/ (read-only)
```

The templates are at:
```
<skills-bash-path>/elca-pptx/ELCAi PPT Template.pptx
<skills-bash-path>/elca-pptx/ELCA PPT Template.pptx
```

In a build script, read the exact path from the system prompt and set:
```python
# Replace <session-name> with the value shown in the system prompt's Shell access section
SKILL_DIR = "/sessions/<session-name>/mnt/.claude/skills/elca-pptx"

# Pick the right template:
TEMPLATE = SKILL_DIR + "/ELCAi PPT Template.pptx"  # for AE / ELCAi content
# TEMPLATE = SKILL_DIR + "/ELCA PPT Template.pptx"  # for everything else

prs = Presentation(TEMPLATE)
ORIG = len(prs.slides)
```

The skills directory is **read-only** from bash — which is fine, `Presentation()` only reads. Saves always go to the outputs directory.

The user sometimes creates a **shell file** — a copy of the template with a title and agenda already filled in. If they provide one, use that instead (it will be in their workspace folder, not the skills dir). Never build slides from scratch with pptxgenjs — always start from a template or shell file.

Both templates use slide dimensions **13.333" × 7.5"** (widescreen 16:9).

---

## Step-by-step workflow

### Step 1 — Think about each slide's message before choosing a layout

This is the most important step and the one most often skipped.

For every slide, ask: **what is the single thing this slide must make the audience understand or feel?** The answer drives the layout and format choice — not the quantity of content.

| If the message is... | Use... |
|---|---|
| A list of parallel facts | Pillars layout |
| A sequence the audience must follow (steps, flow, journey) | Numbered visual boxes — use Blank Slide 1 |
| A dramatic single number or decision | Custom blank with large typographic treatment |
| A comparison between two parties or options | Pillars_2col with aligned rows |
| A question that frames the whole section | Quote/Question layout |
| A process or timeline that has a time axis | Custom blank with drawn timeline |
| A technology stack or layered architecture | Custom blank with horizontal layer bands |

**Do not default to Pillars for everything.** Pillars are the right choice for 3–4 parallel items of equal weight. When content has sequence, drama, or structure that a bullet list flattens, pick a different approach.

Also ask: does this slide need to exist at all? For executive audiences, keep it to **1–2 slides per agenda item**. A slide that duplicates a point made elsewhere should be cut.

### Step 2 — Choose a layout for every slide

Match content type to layout:

| Content type | Layout name | Template |
|---|---|---|
| Opening / closing | `Title Slide`, `Final/Contact Slide` | Both |
| Closing without picture | `Contact Slide without picture` | ELCA only |
| Section break (numbered) | `Chapter Slide 1` | Both |
| Section break (dark/dramatic) | `Chapter Slide Dark lines` | ELCA only |
| 4 stats, 4 categories | `Pillars_4col` | Both |
| 3 items side by side | `Pillars_3col` | Both |
| 2-column comparison, decisions | `Pillars_2col` | Both |
| Running text with sections | `Text Content only 1` | Both |
| Key question / decision point | `Quote/Question 1` | Both |
| Numbered agenda | `Agenda 1` | Both |
| Timeline, architecture, price reveal, stage-setting | `Blank Slide 1` (custom drawing) | Both |

Avoid repeating the same layout on consecutive slides. Vary the rhythm.

### Step 3 — Set up Python

```python
import sys
sys.path.insert(0, '/path/to/skills/elca-pptx/scripts')

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx_helpers import ph, T, LL

prs = Presentation(TEMPLATE_BASH)
ORIG = len(prs.slides)   # number of pre-existing template slides — see Step 5

layouts = {lay.name: lay for lay in prs.slide_master.slide_layouts}
def A(name): return prs.slides.add_slide(layouts[name])
```

### Step 4a — Layout-based slides: use T() and LL() only

```python
# Chapter section break
s = A('Chapter Slide 1')
T(s, 0, '01')
T(s, 1, 'Overview — Current Activities')
T(s, 14, 'Tagline or framing sentence')
# idx 18 = picture placeholder — leave empty

# 4-pillar highlights slide
s = A('Pillars_4col')
T(s, 0, 'Slide Title')
T(s, 1, 'Subtitle — italic, smaller, auto from layout')
T(s, 10, 'Pillar 1 title')   # red, bold, no bullet — from layout
LL(s, 11, [                  # bullets — from layout
    'First bullet point',
    'Second bullet point',
])
T(s, 12, 'Pillar 2 title')
LL(s, 13, ['Bullet A', 'Bullet B'])
T(s, 14, 'Pillar 3 title')
LL(s, 15, ['Item one', 'Item two'])
T(s, 16, 'Pillar 4 title')
LL(s, 17, ['Detail X', 'Detail Y'])
```

Use `''` (empty string) in LL lists as a visual spacer between groups:
```python
LL(s, 11, ['First observation.', 'Second observation.', '',
            'Risk: something to address.', '', '→ Recommended action?'])
```

### Step 4b — Custom blank slides: use TB() and RECT() helpers

For slides that need timelines, architecture diagrams, price reveals, stage-setting with large numbers, or any layout that doesn't fit a Pillars template, use `Blank Slide 1` with custom-drawn shapes.

**Copy these helpers into your build script:**

```python
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

# ── ELCA Primary Colors ───────────────────────────────────────────────────────
RED   = RGBColor(0xE2, 0x36, 0x15)   # ELCA Brickred — titles, accents, CTA boxes
DARK  = RGBColor(0x41, 0x43, 0x44)   # Dark Grey — body text, dark headers
GREY  = RGBColor(0xA0, 0xA0, 0xA0)   # Light Grey — subtitles, secondary labels
LGREY = RGBColor(0xF0, 0xF0, 0xF0)   # Near-white — card / panel backgrounds
WHITE = RGBColor(0xFF, 0xFF, 0xFF)   # White — text on dark/red backgrounds

# ── ELCA Secondary Colors (for multi-category visuals) ────────────────────────
# Use when you need categorical distinction beyond red — diagrams, timelines,
# labeled section boxes. Never let these dominate; RED stays the visual anchor.
# The Sections 1 layout uses exactly this set for its 4-column headers.
BLUE   = RGBColor(0x00, 0xB0, 0xF0)  # Secondary Blue
YELLOW = RGBColor(0xFF, 0xCA, 0x4B)  # Secondary Yellow  (use DARK text on top)
GREEN  = RGBColor(0x28, 0xAF, 0x91)  # Secondary Green

# For 4-category slides use: RED, BLUE, YELLOW, GREEN  (in that order)
# For 3-category slides use: RED, BLUE, GREEN
# For 2-category slides use: RED, DARK

# ── Subtitle coordinates matching Pillars idx=1 (for Blank Slide 1) ──────────
# Title on Blank Slide 1: use T(s, 0, text) — template now has correct position
# Subtitle: no placeholder exists, use BLANKSUB() helper below
_SUB_X, _SUB_Y, _SUB_W, _SUB_H = 0.900, 0.989, 10.211, 0.337


def TB(slide, text, x, y, w, h, size=12, bold=False, italic=False,
       color=DARK, align=PP_ALIGN.LEFT, wrap=True):
    """Floating text box at (x,y) inches. Use for free-standing labels."""
    tx = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tx.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.name = 'Verdana'
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.italic = italic
    r.font.color.rgb = color
    return tx


def RECT(slide, x, y, w, h, fill=None, line_color=None, line_w=0.5):
    """Colored rectangle at (x,y) inches. Returns the shape so you can call SHAPETB on it."""
    shape = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill
    else:
        shape.fill.background()
    if line_color:
        shape.line.color.rgb = line_color
        shape.line.width = Pt(line_w)
    else:
        shape.line.fill.background()
    return shape


def SHAPETB(shape, text, size=11, bold=False, italic=False, color=DARK,
            align=PP_ALIGN.LEFT, valign='middle',
            pad_l=0.08, pad_t=0.06, pad_r=0.08, pad_b=0.06):
    """Put text directly inside a shape's own text frame.

    Use this instead of overlaying a separate TB() on top of a RECT().
    When you move the shape in PowerPoint, the text moves with it.

    This is safe to call add_run() and font.* because the shape was created
    with add_shape() — it has no inherited layout styling to break.
    """
    anchor = {'middle': MSO_ANCHOR.MIDDLE, 'top': MSO_ANCHOR.TOP,
              'bottom': MSO_ANCHOR.BOTTOM}
    tf = shape.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor.get(valign, MSO_ANCHOR.MIDDLE)
    tf.margin_left   = Inches(pad_l)
    tf.margin_top    = Inches(pad_t)
    tf.margin_right  = Inches(pad_r)
    tf.margin_bottom = Inches(pad_b)
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.name  = 'Verdana'
    r.font.size  = Pt(size)
    r.font.bold  = bold
    r.font.italic = italic
    r.font.color.rgb = color
    return shape


def BLANKSUB(slide, text):
    """Subtitle for Blank Slide 1 — same position and style as Pillars idx=1.

    Blank Slide 1 has no subtitle placeholder, so we create a text box at the
    exact same coordinates as the Pillars subtitle (idx=1).
    For the title, use T(s, 0, text) directly — the template's Blank Slide 1
    title placeholder has been repositioned to match Pillars (x=0.653").
    """
    return TB(slide, text, _SUB_X, _SUB_Y, _SUB_W, _SUB_H,
              size=14, italic=True, color=GREY)
```

**Safe content area on Blank Slide 1:**

```
x: 0.65" → 12.85"
y: 1.40" → 6.40"   (content starts below subtitle at ~1.40", footer at ~6.50")
```

Never place content below y=6.40" — the ELCA master logo and footer live there. Never add a red bar at the top or a grey bar at the bottom — the master already provides the ELCA sidebar stripe and bottom logo.

**Minimum font size: 12pt everywhere.** This applies to body text, component labels, captions, and all custom TB() and SHAPETB() calls. Anything smaller is illegible at presentation distance. The templates enforce 12pt as the minimum in all layout `defRPr` elements.

**Typical blank slide structure:**

```python
s = A('Blank Slide 1')
T(s, 0, 'Main Title')          # uses the layout title placeholder — correct position
BLANKSUB(s, 'Subtitle or framing line')
# custom shapes starting at y >= 1.65"
```

**Example: Labeled boxes — use SHAPETB, not overlaid TB**

```python
# WRONG — overlaid text box floats independently of the rectangle
rect = RECT(s, 0.65, 2.20, 1.75, 0.45, fill=RED)
TB(s, 'Medusa.js eShop', 0.68, 2.28, 1.69, 0.29, size=9, bold=True, color=WHITE)

# CORRECT — text lives inside the shape, moves with it
rect = RECT(s, 0.65, 2.20, 1.75, 0.45, fill=RED)
SHAPETB(rect, 'Medusa.js eShop', size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
```

**Example: Stage-setting slide with numbered questions**

```python
s = A('Blank Slide 1')
T(s, 0, 'Der richtige Moment für einen Kurswechsel')
BLANKSUB(s, 'Context subtitle in italic grey')

# Question 01 — number box + text
num = RECT(s, 0.65, 2.20, 0.80, 0.78, fill=RED)
SHAPETB(num, '01', size=26, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
TB(s, 'Key question text?', 1.54, 2.30, 4.85, 0.70, size=12, bold=True, color=DARK)

# Context items — vertical accent bar + free text
RECT(s, 6.80, 1.68, 0.06, 0.55, fill=RED)
TB(s, 'Context Item Title', 6.96, 1.68, 5.65, 0.28, size=12, bold=True, color=DARK)
TB(s, 'Context item description.', 6.96, 1.98, 5.65, 0.35, size=10, color=GREY)
```

**Example: 6-week timeline**

```python
def draw_timeline(slide, y=1.75):
    xl, xr = 0.65, 12.85
    step = (xr - xl) / 6.0
    def wx(n): return xl + n * step

    RECT(slide, xl, y+0.10, xr-xl, 0.06, fill=GREY)
    setup  = RECT(slide, wx(0), y, step,   0.30, fill=RGBColor(0xCC,0xCC,0xCC))
    slice1 = RECT(slide, wx(1), y, step*2, 0.30, fill=RED)
    gap    = RECT(slide, wx(3), y, step,   0.30, fill=RGBColor(0xE0,0xE0,0xE0))
    slice2 = RECT(slide, wx(4), y, step*2, 0.30, fill=DARK)

    SHAPETB(setup,  'Setup',                          size=8, color=DARK)
    SHAPETB(slice1, 'SLICE 1 — Happy Path Order (W1–3)', size=8, bold=True, color=WHITE)
    SHAPETB(gap,    'Bootcamp 2',                     size=8, color=DARK)
    SHAPETB(slice2, 'SLICE 2 — End-to-End + Billing (W4–6)', size=8, bold=True, color=WHITE)

    for i in range(7):
        RECT(slide, wx(i)-0.01, y-0.05, 0.02, 0.09, fill=DARK)
        TB(slide, f'W{i}', wx(i)-0.14, y-0.27, 0.28, 0.22,
           size=8, bold=True, color=DARK, align=PP_ALIGN.CENTER)
    return wx, step
```

**Example: Horizontal architecture layer bands**

```python
layers = [
    ('CHANNEL',  RED,  'eShop (Medusa.js)  ·  .NET Aspire  ·  API Gateway'),
    ('SERVICES', DARK, 'Product Catalog  ·  Eligibility Engine  ·  Order Service'),
    ('ASYNC',    RGBColor(0x1A,0x1A,0x1A), 'Azure Service Bus  ·  Novu  ·  QMC-Mock'),
    ('PLATFORM', RGBColor(0x10,0x10,0x10), 'Azure Container Apps  ·  Azure SQL  ·  Datadog'),
]
ly = 1.50
for lbl, col, content in layers:
    label_box   = RECT(s, 0.65, ly, 1.55, 1.10, fill=col)
    content_box = RECT(s, 2.24, ly, 11.05, 1.10, fill=LGREY,
                       line_color=RGBColor(0xDD,0xDD,0xDD), line_w=0.3)
    SHAPETB(label_box,   lbl,     size=9,  bold=True, color=WHITE,
            align=PP_ALIGN.CENTER, valign='middle')
    SHAPETB(content_box, content, size=11, color=DARK, valign='middle')
    ly += 1.16
```

**Example: Price / investment reveal**

```python
for i, (title, body) in enumerate(value_props):
    bx = 0.65 + i * 4.12
    header = RECT(s, bx, 1.58, 3.92, 0.32, fill=RED)
    card   = RECT(s, bx, 1.90, 3.92, 1.70, fill=LGREY,
                  line_color=RGBColor(0xDD,0xDD,0xDD), line_w=0.3)
    SHAPETB(header, title, size=10, bold=True, color=WHITE, valign='middle')
    SHAPETB(card,   body,  size=10, color=DARK, valign='top', pad_t=0.10)

price_bar = RECT(s, 0.65, 3.65, 12.23, 0.85, fill=DARK)
# For the price bar, two separate elements: label (left) + big number (center)
TB(s, 'FIXPREIS TOTAL', 0.85, 3.70, 3.50, 0.30, size=10, bold=True, color=GREY)
TB(s, "CHF 150'000",   4.50, 3.62, 6.50, 0.88, size=36, bold=True,
   color=RED, align=PP_ALIGN.CENTER)
TB(s, '50% Kickoff  ·  50% nach Abnahme  ·  Gültig bis 20.06.2026',
   0.85, 3.97, 11.90, 0.28, size=9, italic=True, color=WHITE)
```

### Step 5 — Handle pre-existing template slides (critical — avoid zip corruption)

Both templates have pre-existing slides (ELCAi: 4, ELCA: 2). **Do not remove them from `sldIdLst` before adding your slides.** If you do, python-pptx reuses the same filenames for your new slides, creating duplicate entries in the zip that corrupt the PPTX and prevent LibreOffice from opening it.

**Pattern A — fresh build from template:**

```python
# ORIG was set at load time in Step 3: ORIG = len(prs.slides)

# ... add all your slides with A('Layout Name') ...

# Rebuild sldIdLst to include only your slides
all_ids = list(prs.slides._sldIdLst)
our_ids = all_ids[ORIG:]
sldIdLst = prs.slides._sldIdLst
for sid in all_ids: sldIdLst.remove(sid)
for sid in our_ids: sldIdLst.append(sid)

prs.save(OUT)
```

**Pattern B — modifying an existing deck (shell file or previously built PPTX):**

If the PPTX was itself built from a template, it already carries orphaned slides. Saving it again causes new duplicates on top. Use this zip repair after saving:

```python
import zipfile
from lxml import etree
from pptx.oxml.ns import qn

def repair_pptx(src, dst):
    """Remove orphaned slide entries from presentation.xml.rels.
    Keeps only rels that are referenced by sldIdLst; removes duplicates pointing
    to the same slide file that are NOT referenced by sldIdLst."""
    with zipfile.ZipFile(src, 'r') as zin:
        # Build file map keeping LAST occurrence of each filename
        file_data = {}
        for item in zin.infolist():
            file_data[item.filename] = (item, zin.read(item.filename))

        # Find which rIds sldIdLst needs
        prs_xml = etree.fromstring(file_data['ppt/presentation.xml'][1])
        needed_rids = {
            sldId.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
            for sldId in prs_xml.findall('.//' + qn('p:sldIdLst') + '/' + qn('p:sldId'))
        }

        # Fix rels: remove orphan entries pointing to same targets as needed rIds
        rels_item, rels_data = file_data['ppt/_rels/presentation.xml.rels']
        rels_root = etree.fromstring(rels_data)
        needed_targets = {
            rel.get('Target') for rel in rels_root
            if rel.get('Id') in needed_rids and 'slides/slide' in rel.get('Target', '')
        }
        for rel in list(rels_root):
            rid = rel.get('Id')
            tgt = rel.get('Target', '')
            if 'slides/slide' in tgt and 'slideLayout' not in tgt:
                if rid not in needed_rids and tgt in needed_targets:
                    rels_root.remove(rel)
        fixed_rels = etree.tostring(rels_root, xml_declaration=True,
                                    encoding='UTF-8', standalone=True)
        file_data['ppt/_rels/presentation.xml.rels'] = (rels_item, fixed_rels)

    with zipfile.ZipFile(dst, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
        for fname, (item, data) in file_data.items():
            zout.writestr(item, data)
```

**Shell file pattern (title + agenda + final already present):**

If the user provides a shell file with existing content you want to keep, reorder instead of stripping:

```python
sldIdLst = prs.slides._sldIdLst
all_ids = list(sldIdLst)
for sid in all_ids: sldIdLst.remove(sid)
# title(0), agenda(1), your new content(4..n-1), final(3) — skip blank at index 2
for idx in [0, 1] + list(range(4, len(all_ids))) + [3]:
    sldIdLst.append(all_ids[idx])
```

### Step 6 — Save and QA visually

```python
prs.save(OUT)
```

Convert to images and inspect every slide:
```bash
python3 /path/to/skills/pptx/scripts/office/soffice.py --headless --convert-to pdf output.pptx
pdftoppm -jpeg -r 130 output.pdf slide
```

Then `Read` each slide image. Check:
- No text overflows the visible area
- Pillar titles are red (not black)
- Subtitles are italic
- Bullets only where expected
- ELCA logo visible at bottom right (not covered by a custom shape)
- Custom blank slide titles align with titles on layout-based slides (same left edge, same size)
- No floating text boxes sitting on top of colored rectangles — text should be inside the shape

---

## Placeholder index reference

### Pillars (2/3/4 col)

| idx | Role | Layout styling |
|-----|------|---------------|
| 0 | Slide title | Bold, large, left-aligned at x=0.653" |
| 1 | Subtitle | 14pt italic, Light Grey #A0A0A0, no bullet, left-aligned at x=0.653" (matches title) |
| 10 | Pillar 1 title | **Red #E23615, bold, no bullet** |
| 11 | Pillar 1 body | Dark grey, bullet • |
| 12 | Pillar 2 title | Red, bold, no bullet |
| 13 | Pillar 2 body | Dark grey, bullet • |
| 14 | Pillar 3 title | Red, bold, no bullet (Pillars_3col) / KeyMessage (Pillars_2col) |
| 15 | Pillar 3 body | Dark grey, bullet • |
| 16 | Pillar 4 title | Red, bold, no bullet (Pillars_4col) / KeyMessage (Pillars_3col) |
| 17 | Pillar 4 body | Dark grey, bullet • |
| 18 | KeyMessage | Italic, bottom of slide (Pillars_4col) |

**Pattern:** even idx (10,12,14,16) = title (red, no bullet). Odd idx (11,13,15,17) = body (bullets).

### Blank Slide 1

| idx | Role |
|-----|------|
| 0 | Slide title — use `T(s, 0, text)`. Template repositioned to x=0.653", matching all other layouts |
| 27 | Footnote 1 — leave empty |
| 28 | Footnote 2 — leave empty |

No subtitle placeholder exists. Use `BLANKSUB(s, text)` for a subtitle text box at the exact same position and style as `Pillars idx=1`.

### Chapter Slide 1

| idx | Role |
|-----|------|
| 0 | Chapter number (e.g. "01") |
| 1 | Chapter title |
| 14 | Body tagline |
| 18 | Picture — leave empty |

### Text Content only 1

| idx | Role |
|-----|------|
| 0 | Slide title |
| 1 | Main content body (bullets, sections) |
| 16 | Secondary text area |
| 17 | Tertiary text area |

### Agenda 1

| idx | Role |
|-----|------|
| 0 | Title |
| 18 | Chapter names (one per line) |
| 19 | Speakers / descriptions |
| 20 | Times / dates |

### Final/Contact Slide

| idx | Role |
|-----|------|
| 0 | "Thank you" heading (e.g. "Merci.") |
| 1 | Subheading |
| 14 | Primary contact name |
| 15 | Secondary names |
| 16 | Company / job title |
| 17 | Primary email |
| 18 | Website |
| 19 | Picture placeholder — leave empty |

**Note:** this layout is designed for one person. For multiple contacts, use idx 14–18 creatively or replace with a custom Blank Slide 1.

---

## Common issues and fixes

**Pillar title shows black with bullets**
→ You called `p.add_run()` or `text_frame.clear()`. Use `T(slide, idx, text)` only.

**Subtitle not italic / wrong color**
→ Check you're targeting idx=1, not idx=0. `T()` handles `endParaRPr` fallback automatically.

**Text overflows the slide bottom**
→ Too many lines. Reduce bullet count or split into 2 slides.

**Slide order wrong**
→ Use the sldIdLst rebuild pattern in Step 5.

**Layout not found**
→ Check layout names with:
```python
for lay in prs.slide_master.slide_layouts:
    print(lay.name)
```

**LibreOffice can't open the PPTX (zip corruption)**
→ You removed slides from sldIdLst before adding new ones. Use the ORIG pattern from Step 5 instead.

**Custom slide fonts look tiny compared to the rest of the deck**
→ Minimum 11pt for body, 12pt for component labels, 26pt+ for large numbers.

**Grey bar at bottom covers the ELCA logo**
→ You added a rectangle extending past y=6.40". Keep all custom shapes above that line.

**Red bar at top of custom slide looks inconsistent**
→ Remove it. The ELCA master already provides the left sidebar stripe. Adding more chrome creates visual inconsistency across the deck.

**Title on Blank Slide 1 is centered / narrow / wrong position**
→ The template's Blank Slide 1 title placeholder has been repositioned to x=0.653" (matching all other layouts). If you see this issue, you are using an old copy of the template that hasn't been updated. Replace it with the bundled template from the skill directory.

**Text inside a colored box doesn't move when the box is repositioned**
→ You used a floating `TB()` textbox overlaid on a `RECT()`. Use `SHAPETB(rect, text)` instead — it puts the text inside the shape's own text frame so it travels with the shape.

**Subtitle shows a bullet dash "–" and is indented relative to the title**
→ The subtitle placeholder has bullet formatting (no `<a:buNone/>`) and a left indent. Both templates have been fixed: subtitles on all content layouts now have `buNone`, `algn="l"`, x=0.653" (matching the title). If you see this on an old deck, the template was not yet updated.

**Text in a layout or custom shape is smaller than 12pt**
→ Both templates enforce 12pt minimum on all layout `defRPr` elements. For custom `TB()` and `SHAPETB()` calls, always pass `size=12` or larger. Never use `size=9` or `size=10` — those are too small to read on screen.

---

## ELCA brand (hands-off for layout slides — explicit for blank slides)

For layout-based slides, never set colors or fonts explicitly — trust the layout.

For custom blank slides, apply these explicitly:

**Font:** Always **Verdana** for all custom text boxes. Both templates have Verdana set as the theme font (major and minor). Never use Montserrat, Segoe UI, or Calibri in custom shapes.

**Language:** Both templates set `lang="en-US"` **once** in `<p:defaultTextStyle>` (presentation level) and in the master's `<p:txStyles>`. Individual text runs carry **no** explicit `lang` attribute — they inherit en-US from the master default. Do not stamp `lang` on every run; set it once at the master level and let everything fall back cleanly.

**Primary colors:**

| Token | Hex | Use |
|-------|-----|-----|
| `RED   = RGBColor(0xE2, 0x36, 0x15)` | #E23615 | Headers, accents, key numbers, CTA boxes |
| `DARK  = RGBColor(0x41, 0x43, 0x44)` | #414344 | Body text, dark section headers |
| `GREY  = RGBColor(0xA0, 0xA0, 0xA0)` | #A0A0A0 | Subtitles, secondary labels, dividers |
| `LGREY = RGBColor(0xF0, 0xF0, 0xF0)` | #F0F0F0 | Card / panel backgrounds |
| `WHITE = RGBColor(0xFF, 0xFF, 0xFF)` | #FFFFFF | Text on dark/red backgrounds |

**Secondary colors (categorical distinction only):**

Use when a diagram or layout genuinely needs more than red — e.g. 4-category section boxes, labeled swimlanes, role comparison rows. RED stays the visual anchor; secondary colors support it.

| Token | Hex | Use |
|-------|-----|-----|
| `BLUE   = RGBColor(0x00, 0xB0, 0xF0)` | #00B0F0 | 2nd category (e.g. AI, async, tools) |
| `YELLOW = RGBColor(0xFF, 0xCA, 0x4B)` | #FFCA4B | 3rd category — use DARK text on top |
| `GREEN  = RGBColor(0x28, 0xAF, 0x91)` | #28AF91 | 4th category |

For **4 categories**: RED, BLUE, YELLOW, GREEN. For **3**: RED, BLUE, GREEN. For **2**: RED, DARK.

The `Sections 1` template layout uses exactly this set for its column headers.
