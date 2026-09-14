---
name: deck-merger
description: "Build one PowerPoint out of slides cherry-picked from two or more existing .pptx files, copying each slide verbatim so its graphics, SmartArt, embedded objects, fonts and theme survive untouched. Use this skill whenever someone wants to combine, merge, splice, assemble or cut down decks — 'make a deck from the corporate deck plus the BL deck', 'take slides 2-7 and 9-11 from this one and 4-8 from that one', 'strip this deck down to just these slides', 'put together a client deck from our slide library'. Also use it when only one deck is involved but slides need to be removed or reordered, and whenever a previous attempt at rebuilding slides in python-pptx or pptxgenjs lost logos, vector art, diagrams or exact positioning. Use it too when someone wants to browse or pick from what is in their decks — 'show me a catalog of the slides', 'let me choose which slides to include', 'which slides are in these decks' — since it builds an interactive checkbox picker for exactly that."
---

# Deck Merger

Assemble a deck from slides that already exist somewhere else, keeping them
exactly as their author built them.

## Why copy rather than rebuild

The tempting approach — read the source slides, then re-create their content
with python-pptx or pptxgenjs — quietly destroys things: EMF/SVG vector art,
partner logo walls, SmartArt, embedded OLE objects, precise shape positions,
theme inheritance. A client-facing corporate deck is mostly those things. So
this skill copies the slide XML parts verbatim and re-registers them in the
target package instead of regenerating anything. Nothing is retyped, so nothing
drifts.

The other half of that promise is the slide master. A slide's look comes from
its layout, which comes from its master and theme. Merging imported slides onto
the base deck's master would silently restyle them. `merge_decks.py` therefore
imports each source deck's master, layouts and theme alongside the base deck's,
so every slide keeps inheriting what it was designed against. A merged deck
having two or three masters is correct, not a defect.

## Workflow

### Step 1 — Inventory every deck before asking anything

```bash
python3 scripts/inspect_deck.py deck_a.pptx deck_b.pptx
```

This prints, per slide: 1-based position, a `HIDDEN` flag, the layout name and
the opening text. Positions match PowerPoint's slide pane and include hidden
slides — that is the numbering `merge_decks.py` expects, so never renumber it
for the user.

Read the inventory before going back to the user. It answers most of what you
would otherwise ask, and it surfaces the two things that reliably trip up a
merge:

- **Hidden slides.** Decks used as content libraries hide drafts, superseded
  versions and internal-only material. Slides marked "to check", "obsolete" or
  with a yellow banner belong in the same category. Flag them so the user can
  see what they are; leaving them in or out is the user's call, not yours.
- **Housekeeping slides.** Library decks often open with instructions to the
  slide's maintainers ("this deck contains reusable content", "use the Comments
  feature for disclaimers"). Flag these too — they are never client content, but
  say so rather than deciding it silently.

When the deck is visual (org charts, diagrams, photo layouts), the text preview
is thin. Render it and look before describing it:

```bash
python3 <pptx-skill>/scripts/office/soffice.py --headless --convert-to pdf deck.pptx
pdftoppm -jpeg -r 70 deck.pdf slide     # then view slide-*.jpg
```

### Step 2 — Let the user choose the slides

The default way to agree a slide list is the picker: a checkbox catalog of every
deck that sends its own merge instruction back to chat.

```bash
python3 scripts/build_picker.py corporate.pptx business_line.pptx \
    --label "Corp deck 2026" --label "Data & AI" -o picker.html
```

Pass every deck path first, then one `--label` per deck in the same order.
Then pass the file's contents as `widget_code` to the visualizer's `show_widget`
tool. The generated fragment is complete — no title, no explanation inside it;
that prose goes in the chat response around the widget.

- **Every box starts empty, and stays that way.** The selection belongs to the
  user. Do not pre-tick a suggested list, do not tick "the obvious ones", and do
  not treat the badges as a choice made on their behalf. Hidden slides get a
  Hidden badge and housekeeping slides an Internal badge; that is all the
  steering they need.
- `--preselect 1:2-12 --preselect 2:4-13` (1-based deck index) starts from an
  existing selection. Use it only when the user asks to revisit a merge they
  already made, never to open with a recommendation.
- Descriptions are scraped from each slide's own text, so some are cryptic.
  Read them before showing the picker and rewrite the bad ones with
  `--describe 1:13:Untranslated placeholder title`.
- On submit it calls `sendPrompt()` with the selection compacted into ranges,
  which arrives as the user's next message, ready for `merge_decks.py`.

Say what the picker cannot express when handing it over: ascending ranges per
deck only, and deck order is selection order. Anyone who wants a Data & AI slide
as the opening cover has to type that instead.

Describing the decks in prose is the fallback — when the user asks for the
inventory rather than a picker, or when the visualizer is unavailable. Then ask
only what the inventory cannot tell you:

- Which slides from each deck
- Deck order, and whether one deck's title slide should act as a section divider
- Whether they want speaker notes carried over (dropped by default — see below)

Either way, take the answer literally. Do not quietly add slides you think
improve the deck; mention them afterwards instead.

### Step 3 — Merge

```bash
python3 scripts/merge_decks.py -o output.pptx \
    --deck corporate.pptx --slides 2-7,9-11 \
    --deck business_line.pptx --slides 4-8,10
```

- The **first** `--deck` is the base package: its master, theme, size and
  document properties define the output.
- `--slides` takes ranges and single numbers, `all` for everything. Selection
  order is output order, so `--slides 9-11,2` really does put slide 2 last.
- Repeat `--deck`/`--slides` for as many decks as needed; they must come in
  matching pairs.
- `--keep-notes` carries speaker notes across. They are dropped by default
  because notes in library decks are usually written for whoever maintains the
  slide, not for the audience, and they travel with author metadata.

The script prunes everything the selection left unreferenced, so the output is
usually far smaller than the sum of its inputs.

### Step 4 — Verify, then look

Validate the package first — it catches the structural faults PowerPoint
refuses to open:

```bash
python3 <pptx-skill>/scripts/office/validate.py output.pptx --original corporate.pptx
```

Pass `--original` pointing at the base deck so faults the source already had
don't read as yours.

Then render and view every slide (same two commands as Step 1). Check slide
order, that no slide arrived blank, and that logos and diagrams are present.

### Step 5 — Hand it over

State the final slide order, note anything you left out and why, and say which
deck each half came from.

## Reading render artifacts correctly

Renders go through LibreOffice, which is not PowerPoint. It substitutes missing
fonts and rasterizes EMF/WMF vector art badly. Both produce convincing-looking
damage — text overlapping a shape, mangled logo lettering — on slides that are
perfectly fine in PowerPoint. This costs real time if you take it at face value.

Before reporting a visual defect, or acting on a user's report of one, render
the *source* deck at the same slide. If it looks equally broken there, the merge
didn't cause it, and the odds are good that nothing is broken at all.

To settle it properly, compare the slide XML directly. Merged slides should be
byte-identical to their source except for relationship IDs:

```python
import re, zipfile, hashlib
def norm(b):   # rIds are renumbered on import; everything else must match
    return hashlib.md5(re.sub(rb'r:(id|embed|link)="rId\d+"', b'', b)).hexdigest()
```

If the bytes match and every referenced part is present at the same size, the
merge is faithful and the problem is the renderer or the source. Say so plainly
rather than "fixing" a slide that isn't broken.

## Failure modes worth knowing

**Duplicate layout IDs.** Two masters in one package must not share
`sldLayoutId` values or PowerPoint reports the file as corrupt. The script
renumbers on import; if you ever hand-merge masters, do the same.

**Orphaned comment parts.** Decks with modern comments often carry
`ppt/comments/modernComment_*.xml` files that survive slide deletion and then
fail validation as unreferenced. The script's prune step removes them along with
their content-type overrides.

**A slide that looks empty after merging** usually means its layout or master
didn't come across. Check the slide's `.rels` for a `slideLayout` target that
exists in the output package.

**Charts and embedded workbooks** are copied as parts, so they keep working, but
a duplicated chart shares its data part with the original. Editing one edits
both — split them if the user needs divergent numbers.

## Scripts

| Script | Purpose |
|---|---|
| `scripts/inspect_deck.py` | Numbered inventory of one or more decks, with hidden-slide flags |
| `scripts/build_picker.py` | HTML for a checkbox slide picker to pass to `show_widget`; `-h` for full usage |
| `scripts/merge_decks.py` | The merge itself; `-h` for full usage |

`merge_decks.py` needs only the standard library; the other two also need
`python-pptx`. Validation and rendering use the `pptx` skill's
`scripts/office/` helpers when available.
