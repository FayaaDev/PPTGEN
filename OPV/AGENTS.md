# OPV presentations

## File roles

- `generate.cjs`: reusable Compact layout definitions and template/example generator. Uses the local PptxGenJS build (`../dist/pptxgen.cjs.js`) and JSZip. Reads two illustrations from `Polio_2026_7.pptx`.
- `OPV_Compact_Master.potx`: generated PowerPoint template for fresh presentations. Open it in PowerPoint and choose layouts under Home → New Slide.
- `OPV_Compact_Examples.pptx`: generated 12-slide layout examples; created by `generate.cjs`, may not exist until generation.
- `recreate.py`: conversion script for the existing 61-slide Polio deck. Slide numbers, layout mapping, and special positioning are specific to that deck; do not use it as a generic fresh-presentation generator.
- `verify.py`: structural checks for the generated template and examples.
- `Polio_2026_7_OPV_Compact_checks.json`: generated preservation/verification report, not configuration.
- `Polio_2026_7.pptx` and `Polio_2026_7.pdf`: original source/reference. Preserve them.
- `Polio_2026_7_OPV_Compact.pptx` and its PDF: recreated editable deck and rendered preview.
- Files containing `Repaired` are historical artifacts, not delivery files.

## Design

- Widescreen 16:9, approximately 13.333 × 7.5 inches.
- Font: Franklin Gothic Book.
- Palette: ink `0C2A33`, teal `136F7A`, blue `2B6F94`, green `2D7A62`; pale rectangular cards and teal accent rules.
- Twelve layouts: Cover, Section, Content, Comparison, Three cards, Table, Chart, Image, Timeline, Summary, Appendix, References.
- Use matching solid edge frames: dark ink on white slides, white on dark slides. Thickness is 0.16 inches (11.52 pt), centered 0.08 inches inward so the stroke reaches the slide edge.
- Keep text, tables, and charts editable. Preserve image aspect ratios. Keep main content above the citation/footer zone.

## Current implementation gap

The matching thick frames are currently implemented in `recreate.py` as slide-level shapes. They are not yet in `generate.cjs` or the master template. Before generating a fresh presentation that must match the finished deck, add the frames to the reusable layout definitions and regenerate/verify the template. Account for the full-slide frames in `verify.py`'s citation-zone check: decorative borders are not body content.

## Workflow

Commands below run from the project root:

```sh
# Generate the Compact template and examples
node OPV/generate.cjs
python3 OPV/verify.py

# Rebuild and verify the specific 61-slide Polio deck
python3 OPV/recreate.py

# Verify the existing recreated deck without rebuilding
python3 OPV/recreate.py --verify-only
```

- For fresh decks generated with code, reuse the exported `createOpvPresentation()` and `layouts` from `generate.cjs`.
- Change reusable design in `generate.cjs`; keep Polio-specific conversion adjustments in `recreate.py`.
- Preserve source text order, media, chart/workbook payloads, notes, and relationships when converting the existing deck.
- Keep Office master/layout IDs globally unique and shape IDs unique within each XML part.
- Structural checks do not prove visual fit. Open changed output in PowerPoint and inspect rendered slides for clipping, overlap, and readability.
- PDF export is a separate PowerPoint step. Re-export after deck changes; scripts do not generate PDFs.
- For AppleScript PDF export, use a POSIX file object: `save active presentation in (POSIX file "/absolute/path/output.pdf") as save as PDF`.
- Do not overwrite or remove unfamiliar presentations or repaired artifacts without investigating their ownership.
