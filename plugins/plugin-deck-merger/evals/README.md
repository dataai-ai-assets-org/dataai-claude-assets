# Evals for plugin-deck-merger

Cases here exercise all three bundled skills end to end (real PPTX
generation/merging), so they need file-writing tools that `claude plugin
eval` gates behind an explicit operator grant:

```bash
claude plugin eval . --allow-tools Bash,Write --trust-plugin
```

Run a single case while iterating:

```bash
claude plugin eval . --case deck-merger-basic --runs 1 --ablation none --allow-tools Bash,Write
```

`deck-merger-basic` ships its own fixture decks under
`deck-merger-basic/fixtures/` (generated with python-pptx — see git history)
so the case is self-contained and reproducible.
