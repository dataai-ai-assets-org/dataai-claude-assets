---
type: llm
weight: 1
---

PASS if the final response indicates that `output.pptx` was created with exactly 3 slides — a title slide, a "Why Modernize" content slide with bullet points, and a closing "Thank You" slide — built from the plain **ELCA** template (not ELCAi, since the topic is not Agentic Engineering).

FAIL if the agent reports an error, used the ELCAi template for this non-Agentic-Engineering topic, produced the wrong number of slides, or did not go through the elca-pptx skill's helpers.
