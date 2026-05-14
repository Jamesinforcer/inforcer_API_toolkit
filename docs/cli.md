# Command-Line Flags

The interactive menu covers the common cases. For scripted or repeat runs, you can invoke the generator directly.

## Phase 1 — export selection workbooks

```bash
python inforcer_sop_generator.py --export-selection
python inforcer_sop_generator.py --export-selection --tenant "Contoso"
python inforcer_sop_generator.py --export-selection --all
```

## Phase 3 — build an SOP from a ticked workbook

```bash
python inforcer_sop_generator.py --from-selection "output/Contoso_Policy_Selection.xlsx"
```

## Legacy one-shot auto-match

Skips the selection step — every policy is considered for Part 2.

```bash
python inforcer_sop_generator.py --auto-match --tenant "Contoso"
```

## Global flags

| Flag | Default | Description |
|------|---------|-------------|
| `--region {anz,eu,uk,us}` | `anz` | Inforcer region endpoint |
| `--out DIR` | `./output` | Output directory |
| `--template PATH` | `MSP-M365-Baseline-SOP.docx` | Path to the SOP template |
