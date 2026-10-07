# SABIC presentation template

Based on the visual design of `Template.pdf`: original logo and gradient ribbons,
white backgrounds, grey headings, cyan financial figures and horizontal table rules.
Slide size matches the PDF: **11.81 × 7.87 inches (approximately 3:2)**.

## Open and use

- **SABIC-Template.potx** — reusable PowerPoint template; open to create a new deck.
- **SABIC-Template.pptx** — eight editable example slides; duplicate and replace content.
- **SABIC-Template-preview.pdf** — visual preview exported using Microsoft PowerPoint.

Included layouts: cover/section divider, single column, two columns, three columns,
financial table, and text with table. Choose them under **Home → New Slide**.
Tables and text are native editable PowerPoint objects. Logo and ribbon artwork
are images on the layouts; change them under **View → Slide Master**.

This is a reusable design template, not a conversion of all 23 financial-report pages.
Example copy and financial dashes are placeholders for your own content.

## Regenerate

From the repository root:

```sh
npm install --omit=dev --ignore-scripts --package-lock=false
node SABIC/generate.cjs
python3 SABIC/verify.py
```

Edit `generate.cjs` for slide content; edit `template.cjs` for shared layouts.
Regeneration writes the PPTX and POTX. Export a fresh PDF from PowerPoint after changes.

The supplied PDF embeds subsets of SABIC's proprietary typefaces. The template uses
**Helvetica Neue** as a fallback; fonts are not embedded. For the original typography,
install the full licensed SABIC fonts and regenerate with their installed family names:

```sh
SABIC_BODY_FONT="SABIC Typeface Text" \
SABIC_HEADING_FONT="SABIC Typeface Headline" \
node SABIC/generate.cjs
```

Exact installed font family names may differ. Match them to your system's font list.

## Reuse from JavaScript

```js
const { createSabicPresentation } = require('./SABIC/template.cjs');
const pptx = createSabicPresentation();
const slide = pptx.addSlide({ masterName: 'SABIC_TWO_COLUMNS' });
slide.addText('BUSINESS UPDATE', { placeholder: 'title' });
slide.addText('Reporting period', { placeholder: 'subtitle' });
slide.addText('PERFORMANCE', { placeholder: 'leftHeading' });
slide.addText('Your performance summary', { placeholder: 'left' });
slide.addText('NEXT STEPS', { placeholder: 'rightHeading' });
slide.addText('Your planned actions', { placeholder: 'right' });
pptx.writeFile({ fileName: 'business-update.pptx' });
```

Original assets are included. To re-extract them from this specific PDF, install
`pypdf` and run `python3 SABIC/extract-assets.py`. Extraction preserves ribbon transparency.
