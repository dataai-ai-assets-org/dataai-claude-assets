# plugin-deck-merger

PowerPoint merging toolkit for ELCA's Data, Analytics & AI Business Line. Assembles new decks from slides cherry-picked across existing presentations while preserving all graphics, formatting, themes, and embedded objects verbatim.

## What it does

The **deck-merger** skill combines, reorders, and filters slides from multiple PowerPoint files into a single output deck. Unlike rebuild-based approaches, it copies slide XML parts directly—preserving vector art, SmartArt, partner logos, embedded objects, and exact positioning that would be lost if slides were recreated with python-pptx.

| Task | Use this skill |
|---|---|
| Merge two or more decks into one | "Combine the corporate deck with the BL deck" |
| Pick specific slides from decks | "Take slides 2-7 and 9-11 from this one" |
| Reorder or trim a deck | "Strip this deck down to just these slides" |
| Browse a deck's contents | "Show me what's in these presentations" |
| Assemble a client deck from slide library | "Make a deck from our slide library" |

## How it works

**Step 1: Inventory** — scan each source deck to see what slides are available
```bash
python3 scripts/inspect_deck.py deck_a.pptx deck_b.pptx
```

**Step 2: Choose** — use an interactive checkbox picker to select slides (or describe them in prose)
```bash
python3 scripts/build_picker.py corporate.pptx business_line.pptx \
    --label "Corp 2026" --label "Data & AI" -o picker.html
```

**Step 3: Merge** — assemble the output deck with exact slide order
```bash
python3 scripts/merge_decks.py -o output.pptx \
    --deck corporate.pptx --slides 2-7,9-11 \
    --deck business_line.pptx --slides 4-8,10
```

**Step 4: Validate & inspect** — check the result and hand it over
```bash
python3 <pptx-skill>/scripts/office/validate.py output.pptx --original corporate.pptx
```

## Why copy instead of rebuild

Rebuilding slides with python-pptx or pptxgenjs silently destroys:
- Vector art (EMF/SVG) and rasterized graphics
- SmartArt and diagrams
- Embedded OLE objects (Excel workbooks, Visio charts)
- Precise shape positions and alignments
- Theme and master inheritance

A client-facing corporate deck is mostly these things. **Deck merger copies the slide XML verbatim**, so nothing drifts and nothing breaks. Each imported slide keeps inheriting from its original master, layouts, and theme—merged decks with multiple masters is correct, not a defect.

## Layout

```
plugin-deck-merger/
├── .claude-plugin/
│   └── plugin.json                      # manifest
├── README.md                            # this file
└── skills/
    └── deck-merger/
        ├── SKILL.md                     # detailed workflow and usage
        └── scripts/
            ├── inspect_deck.py          # inventory one or more decks
            ├── build_picker.py          # interactive slide picker (HTML)
            └── merge_decks.py           # merge decks by slide range
```

## Dependencies

- **Python 3.7+**
- **python-pptx** (for `inspect_deck.py` and `build_picker.py`)
- Standard library only for `merge_decks.py`
- Optional: LibreOffice (for validation and rendering via `pptx-skill`)

## Install & test

Local test run, from the directory containing the plugin folder:

```bash
claude --plugin-dir ./plugin-deck-merger
```

Then use `/reload-plugins` to pick up edits without restarting. To distribute it to the BL, publish through a plugin marketplace repository; see https://code.claude.com/docs/en/plugin-marketplaces.

## Key workflows

### Inventory before you merge
Always run `inspect_deck.py` first. It answers most questions and surfaces:
- **Hidden slides** (drafts, internal-only material, obsolete versions)
- **Housekeeping slides** (library instructions, maintainer notes)
- Slide positions and text preview (for matching the inventory to user intent)

### Let users choose via picker
The interactive picker (`build_picker.py`) is the default UX:
- Users see every slide with its text and badges (Hidden, Internal)
- Every box starts empty—the selection belongs to the user
- Selection is returned as a compact merge instruction ready for `merge_decks.py`

Fall back to prose only when the user asks for the inventory instead, or when the visualizer is unavailable.

### Trust the merge
Merged slides are byte-identical to their source (except for relationship IDs). If the rendering looks wrong, compare the source deck's render at the same slide first—LibreOffice substitutes fonts and rasterizes vector art badly, which produces convincing-looking but harmless damage.

## Troubleshooting

| Symptom | Likely cause | Fix |
|---|---|---|
| Slide looks blank after merge | Layout or master didn't import | Check slide `.rels` for valid `slideLayout` target |
| PowerPoint reports file corrupt | Duplicate layout IDs across masters | The script renumbers on import (should not happen) |
| Orphaned comment parts fail validation | Slides with modern comments left dangling parts | `merge_decks.py` prunes these automatically |
| Chart or workbook data edits both | Duplicated charts share data parts | Split them if user needs divergent numbers |

## Versioning

Versioning is by commit SHA (resolves from marketplace repo). The plugin.json intentionally has no `version` field—users get an update whenever you push. If you add a `version`, it pins the plugin and you must bump it on every release.

## Related

- **elca-pptx** skill: For rebuilding slides or creating new ones with the ELCA template
- **pptx-skill**: Shared LibreOffice validation and rendering utilities
