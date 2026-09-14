---
type: llm
weight: 1
---

PASS if the response indicates it inventoried both source decks (e.g. via `inspect_deck.py`) before merging, and produced `merged.pptx` containing exactly the 4 requested slides (2 from `deck-a.pptx`, 3 from `deck-b.pptx`, in the requested order) by copying the original slide XML rather than rebuilding the slides with python-pptx/pptxgenjs.

FAIL if it rebuilt slide content from scratch instead of copying, produced the wrong slide count or order, or skipped inventorying the source decks first.
