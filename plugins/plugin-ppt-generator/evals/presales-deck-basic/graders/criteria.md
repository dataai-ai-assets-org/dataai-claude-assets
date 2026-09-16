---
type: llm
weight: 1
---

PASS if the response describes generating a single PPTX deck built from the content library that includes: general ELCA introduction slides, Data & AI BL capability slides, a logistics-relevant industry module (or an explicit note that none exists for logistics), and an AI-use-case slide drafted specifically for Northwind Freight / the logistics industry (not generic boilerplate). Since the prompt explicitly said there's no reference deck for this call, the response should say plainly that the deck has no reference-project slides this time rather than fabricating any.

FAIL if the deck content is generic with nothing specific to Northwind Freight or logistics, the run errors out, the presales-deck-generator skill was not used, or the response presents fabricated reference-project content despite none being supplied.
