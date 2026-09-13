# plugin-ppt-generator

PowerPoint tooling for ELCA's Data, Analytics & AI Business Line. Bundles the
branded rendering toolkit and the pre-sales deck generator that builds on it,
so they install, version, and update as one unit.

## What's inside

| Skill | What it does | Triggers on |
|---|---|---|
| `elca-pptx` | Builds natively ELCA-branded PPTX from the ELCA / ELCAi templates — correct fonts, colors, bullets — using the `T()` / `LL()` helpers that avoid overriding layout styling. | Any request to create, rebuild, or add slides to an ELCA-looking deck; also fixing decks with wrong colors, missing bullets, or overridden fonts. |
| `presales-deck-generator` | Assembles a first-call pre-sales deck for a named prospect: general ELCA slides, Data & AI BU slides, an industry module, a company-specific AI use-case slide grounded in public research, and 4–5 reference projects. | "Build me a deck for our call with Acme AG", "slides for the first meeting with …", regenerating an existing pre-sales deck for a different company. |

`presales-deck-generator` renders through `elca-pptx` rather than drawing
slides itself. That dependency is the reason these two travel together:
installed separately, the generator has to be told where the template and
helpers live; installed as this plugin, it finds them as a sibling under
`skills/` with no configuration.

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
            ├── content_library.py      # the maintained, non-per-call content
            └── build_presales_deck.py  # assembles the deck
```

Component directories sit at the plugin root, never inside `.claude-plugin/`
— putting them there loads the plugin but silently hides the skills.

## Install

Local test run, from the directory containing the plugin folder:

```bash
claude --plugin-dir ./plugin-ppt-generator
```

Then `/reload-plugins` to pick up edits without restarting. To distribute it
to the BU, publish it through a plugin marketplace repository; see
https://code.claude.com/docs/en/plugin-marketplaces.

## Maintenance

- **Content changes** (ELCA facts, BU stats, industry modules, reference
  projects) go in `skills/presales-deck-generator/scripts/content_library.py`.
  That file is plain, commented Python and is meant to be edited by BU
  champions, not only by engineers.
- **Versioning is by commit.** `plugin.json` deliberately has no `version` field,
  so Claude Code resolves the version from the marketplace repo's commit SHA and
  users get an update whenever you push. If you ever add a `version`, it pins the
  plugin and you must bump it on every release or installed users keep the old
  copy.
- Never add a reference project that doesn't correspond to a real engagement,
  and never let the generator invent an AI use case it can't source. Both
  rules are explained where they bite, in `presales-deck-generator/SKILL.md`.
