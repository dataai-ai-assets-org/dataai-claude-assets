# dataai-claude-assets

Claude Code plugins and customization assets for ELCA's Data, Analytics & AI Business Line.

## Plugin

### plugin-deck-merger

Assembles new decks from slides cherry-picked across existing presentations. Copies slide XML verbatim so graphics, SmartArt, embedded objects, fonts and themes survive untouched.

Key features:
- Inventory existing decks to see what slides are available
- Interactive checkbox picker to choose which slides to include
- Merge multiple decks by slide range, preserving all formatting
- Validate and render the result

**See**: [plugin-deck-merger/README.md](plugins/plugin-deck-merger/README.md)

## Marketplace

The top-level `.claude-plugin/marketplace.json` defines how these plugins are discovered and listed when published to a Claude Code plugin marketplace. Update the `owner` field with the BL contact email and version information as needed.

## Structure

```
dataai-claude-assets/
├── .claude-plugin/
│   └── marketplace.json                 # marketplace metadata
├── README.md                            # this file
└── plugins/
    └── plugin-deck-merger/              # deck merge & pick tool
        ├── .claude-plugin/plugin.json
        ├── README.md
        └── skills/
            └── deck-merger/             # merge, reorder, filter slides
                ├── SKILL.md
                ├── README.md
                └── scripts/
                    ├── inspect_deck.py
                    ├── build_picker.py
                    └── merge_decks.py
```

## Quick start

### Load the plugin in Claude Desktop

For local testing:
```bash
claude --plugin-dir ./plugins/plugin-deck-merger
```

Then use `/reload-plugins` to pick up script edits without restarting.

### Publish to marketplace

1. Ensure each plugin's metadata is correct in `.claude-plugin/plugin.json`
2. Commit and push to your plugin marketplace repository
3. Plugins are discovered by commit SHA (no manual versioning required)

See https://code.claude.com/docs/en/plugin-marketplaces for full distribution docs.

## Maintenance

- **Skill-specific content** (when a skill maintains any) goes in the `scripts/` folders under each skill — presales-deck-generator currently maintains none; see its SKILL.md
- **Skill logic** lives in `.md` files alongside the scripts
- **Plugin metadata** is in `.claude-plugin/plugin.json`
- **Marketplace configuration** is in `.claude-plugin/marketplace.json` at the repo root

## License

UNLICENSED — internal ELCA assets only.

## Contact

Data, Analytics & AI Business Line  
(Update `marketplace.json` with primary contact email)
