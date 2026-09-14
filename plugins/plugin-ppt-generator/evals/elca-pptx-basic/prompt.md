---
name: elca-pptx-basic
tags: [smoke, elca-pptx]
runs: 1
max_turns: 15
timeout_seconds: 300
allowed_tools: [Read, Glob, Grep, Skill]
---

Build a 3-slide ELCA-branded PowerPoint deck for an internal Sounding Board and save it as `output.pptx` in the current directory:

1. A title slide: title "Sounding Board: Data Platform Modernization", subtitle "ELCA Data, Analytics & AI".
2. A content slide titled "Why Modernize" with 3 bullet points of your choosing.
3. A closing "Thank You" slide.

This deck is about a general ELCA offer, not about Agentic Engineering, ELCAi, BMAD, or Claude Code.
