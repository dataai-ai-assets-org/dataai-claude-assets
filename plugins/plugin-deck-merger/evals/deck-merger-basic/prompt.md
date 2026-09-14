---
max_turns: 15
timeout_seconds: 300
allowed_tools: [Read, Glob, Grep, Skill]
---

I have two decks in the `fixtures/` directory: `deck-a.pptx` (an "Acme Corp" overview, 3 slides) and `deck-b.pptx` (ELCA reference projects, 4 slides). Build one merged deck called `merged.pptx` in the current directory containing, in this order:

1. Slide 1 from `deck-a.pptx` (the "Acme Corp Overview" title slide)
2. Slide 2 from `deck-a.pptx` ("Acme Corp: Key Facts")
3. Slides 2, 3 and 4 from `deck-b.pptx` (the three reference project slides)

Keep every slide's original formatting exactly as it appears in its source deck — don't rebuild the slides from scratch.
