# plugin-ppt-generator

PowerPoint tooling for ELCA's Data, Analytics & AI Business Line. Bundles the
branded rendering toolkit and the pre-sales deck generator that builds on it,
so they install, version, and update as one unit.

## What's inside

| Skill | What it does | Triggers on |
|---|---|---|
| `elca-pptx` | Builds natively ELCA-branded PPTX from the ELCA / ELCAi templates — correct fonts, colors, bullets — using the `T()` / `LL()` helpers that avoid overriding layout styling. | Any request to create, rebuild, or add slides to an ELCA-looking deck; also fixing decks with wrong colors, missing bullets, or overridden fonts. |
| `presales-deck-generator` | Assembles a first-call pre-sales deck for a named prospect: an industry-opportunity slide and a company-specific AI use-case slide, both drafted live from public research, plus reference-project slides merged in verbatim from a reference deck the user points to — no general ELCA or Business Line content, no maintained content library at all, just what's specific to the call. | "Build me a deck for our call with Acme AG", "slides for the first meeting with …", regenerating an existing pre-sales deck for a different company. |

`presales-deck-generator` renders through `elca-pptx` rather than drawing
slides itself. That dependency is the reason these two travel together:
installed separately, the generator has to be told where the template and
helpers live; installed as this plugin, it finds them as a sibling under
`skills/` with no configuration.

It also depends on the **deck-merger** skill (in the separate
**plugin-deck-merger** plugin) to splice in reference-project slides — that
one isn't bundled here, so install both plugins to get the full workflow. See
"Dependency: deck-merger" in `presales-deck-generator/SKILL.md`.

## Layout

```
plugin-ppt-generator/
├── .claude-plugin/
│   └── plugin.json                     # manifest — only this file lives here
├── README.md
└── skills/
    ├── elca-pptx/
    │   ├── SKILL.md
    │   ├── ELCA PPT Template.pptx      # corporate template
    │   ├── ELCAi PPT Template.pptx     # ELCAi template
    │   └── scripts/pptx_helpers.py     # T, LL, SHAPETB, BLANKTITLE
    └── presales-deck-generator/
        ├── SKILL.md
        └── scripts/
            └── build_presales_deck.py  # assembles the deck from per-call content
```

Component directories sit at the plugin root, never inside `.claude-plugin/`
— putting them there loads the plugin but silently hides the skills.

## Install

Local test run, from the directory containing the plugin folder:

```bash
claude --plugin-dir ./plugin-ppt-generator
```

Then `/reload-plugins` to pick up edits without restarting. To distribute it
to the BL, publish it through a plugin marketplace repository; see
https://code.claude.com/docs/en/plugin-marketplaces.

## Maintenance

- **There is no maintained content library.** `presales-deck-generator` has
  no general ELCA content, no Business Line content, and no industry-module
  dict — everything (industry angle, AI use cases, references) is drafted
  or sourced fresh per call. See "Why this skill is shaped the way it is" in
  `presales-deck-generator/SKILL.md`.
- **Versioning is by commit.** `plugin.json` deliberately has no `version` field,
  so Claude Code resolves the version from the marketplace repo's commit SHA and
  users get an update whenever you push. If you ever add a `version`, it pins the
  plugin and you must bump it on every release or installed users keep the old
  copy.
- Never present a reference slide that isn't a real slide copied from the
  user-supplied reference deck, and never let the generator invent an AI use
  case it can't source. Both rules are explained where they bite, in
  `presales-deck-generator/SKILL.md`.
