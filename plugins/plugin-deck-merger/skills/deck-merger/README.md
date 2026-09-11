# Deck Merger Skill

Merge, reorder, and cherry-pick slides from multiple PowerPoint files into a single output deck while preserving all graphics, formatting, themes, and embedded objects.

## Scripts

### `inspect_deck.py`
Inventory the contents of one or more decks without modifying them.

```bash
python3 scripts/inspect_deck.py deck_a.pptx deck_b.pptx
```

Output shows per slide:
- 1-based position (matching PowerPoint slide pane)
- `HIDDEN` flag (if slide is hidden)
- Layout name
- Opening text (first words on the slide)

**Use this first.** It answers most questions and flags hidden slides and library housekeeping material before you decide what to merge.

### `build_picker.py`
Generate an interactive HTML checkbox picker for selecting slides.

```bash
python3 scripts/build_picker.py deck_a.pptx deck_b.pptx \
    --label "Corporate 2026" --label "Data & AI" \
    -o picker.html
```

Options:
- `--label` (one per deck, in order) — display name for each deck
- `-o, --output` — output HTML file path
- `--preselect 1:2-7 --preselect 2:4-8` — pre-tick slides (deck index : slide ranges)
- `--describe 1:5:Custom title` — override auto-scraped slide text

**Use this as the default UX** when the user wants to select slides. The output is complete HTML ready to pass to the visualizer's `show_widget` tool. On submit, it returns a selection compacted into ranges.

### `merge_decks.py`
Perform the merge itself.

```bash
python3 scripts/merge_decks.py -o output.pptx \
    --deck base.pptx --slides all \
    --deck source_a.pptx --slides 2-7,9-11,13 \
    --deck source_b.pptx --slides 4-8
```

Arguments:
- `-o, --output` — output file path
- `--deck` — path to a source deck (first one is the base; its master, theme, size and document properties define the output)
- `--slides` — which slides to take from this deck (ranges like `2-7,9-11`, single numbers, or `all`)
- `--keep-notes` — carry speaker notes across (dropped by default)

**Selection order is output order**, so `--slides 9-11,2-7` puts slides 2-7 last.

**Repeat `--deck` and `--slides` in matching pairs** for as many source decks as you need.

The script only requires Python's standard library. The other scripts also need `python-pptx`.

## Workflow

1. **Inventory**: Run `inspect_deck.py` on every deck to see what you're working with
2. **Choose**: Use `build_picker.py` to generate an interactive picker (or ask the user to describe their selection in prose)
3. **Merge**: Run `merge_decks.py` with the selected ranges
4. **Validate**: Run validation through the pptx-skill's `scripts/office/validate.py` to catch structural faults
5. **Render**: Optional—check the result visually through LibreOffice render (for spot-checking, not final approval)
6. **Hand over**: State the final slide order, note what was left out and why, and say which deck each slide came from

## Key points

- **Merged slides are byte-identical to their source** (except for relationship IDs), so nothing is lost—text, graphics, SmartArt, embedded objects, positioning, fonts all survive
- **Multiple masters is correct**—each imported slide keeps inheriting from its original master, layouts and theme
- **Hidden slides and housekeeping are real**—flag them so users can choose, don't decide silently
- **LibreOffice rendering is not gospel**—it substitutes fonts and rasterizes vector art badly; compare the source deck's render to settle visual questions
- **Duplicate layout IDs are corruption**—the script prevents this; if you ever hand-merge masters, renumber layout IDs yourself

## Dependencies

- **Python 3.7+**
- **python-pptx** (for `inspect_deck.py` and `build_picker.py`)
- **Standard library only** for `merge_decks.py`
- **LibreOffice** (optional, for validation and rendering)

## Full documentation

See [SKILL.md](SKILL.md) for complete workflow details, troubleshooting, and render artifact interpretation.
