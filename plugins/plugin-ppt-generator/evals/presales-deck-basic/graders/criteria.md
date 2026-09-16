---
type: llm
weight: 1
---

PASS if the response describes generating a single PPTX deck that includes: a logistics-relevant industry module (or an explicit note that none exists for logistics), and an AI-use-case slide drafted specifically for Northwind Freight / the logistics industry (not generic boilerplate). The deck should NOT include general ELCA corporate slides or Data & AI Business Line capability slides — this skill no longer produces those. Since the prompt explicitly said there's no reference deck for this call, the response should say plainly that the deck has no reference-project slides this time rather than fabricating any.

FAIL if the deck content is generic with nothing specific to Northwind Freight or logistics, the run errors out, the presales-deck-generator skill was not used, or the response presents fabricated reference-project content despite none being supplied.
