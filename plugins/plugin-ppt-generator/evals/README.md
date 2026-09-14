# Evals for plugin-ppt-generator

Cases here exercise both bundled skills end to end (real PPTX generation), so
they need file-writing tools that `claude plugin eval` gates behind an
explicit operator grant:

```bash
claude plugin eval . --allow-tools Bash,Write --trust-plugin
```

Run a single case while iterating:

```bash
claude plugin eval . --case elca-pptx-basic --runs 1 --ablation none --allow-tools Bash,Write
```
