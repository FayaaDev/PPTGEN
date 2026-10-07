const path = require('node:path');
const PptxGenJS = require('../dist/pptxgen.cjs.js');

const COLORS = { blue: '0047BA', cyan: '009FCC', yellow: 'FFAA00', text: '575756', muted: '888888', rule: '999999' };
const SIZE = { width: 850.32 / 72, height: 566.64 / 72 };

function createSabicPresentation({ bodyFont = 'Helvetica Neue', headingFont = bodyFont } = {}) {
	const pptx = new PptxGenJS();
	pptx.defineLayout({ name: 'SABIC', ...SIZE });
	pptx.layout = 'SABIC';
	pptx.author = 'SABIC';
	pptx.subject = 'Reusable layouts based on SABIC/Template.pdf';
	pptx.title = 'SABIC Presentation Template';
	pptx.company = 'Saudi Basic Industries Corporation';
	pptx.lang = 'en-GB';
	pptx.theme = { headFontFace: headingFont, bodyFontFace: bodyFont, lang: 'en-GB' };
	const text = { fontFace: bodyFont, fontSize: 12, color: COLORS.text, margin: 0, align: 'left', valign: 'top' };
	const image = (name, x, y, w, h) => ({ image: { path: path.join(__dirname, `${name}.png`), x, y, w, h } });
	const placeholder = (name, type, x, y, w, h, options = {}) => ({
		placeholder: { options: { ...text, name, type, x, y, w, h, ...options }, text: '' },
	});
	pptx.defineSlideMaster({
		title: 'SABIC_COVER', background: { color: 'FFFFFF' },
		objects: [
			image('logo-yellow', 673.34 / 72, 63.15 / 72, 133.92 / 72, 69.98 / 72),
			image('ribbon-yellow', 680.55 / 72, 397.72 / 72, 169.77 / 72, 164 / 72),
			image('ribbon-blue', 331.85 / 72, 271.76 / 72, 518.47 / 72, 288.5 / 72),
			{ text: { text: 'CHEMISTRY THAT MATTERS™', options: { ...text, x: 0.62, y: 0.88, w: 4.5, h: 0.25, fontSize: 13 } } },
			placeholder('title', 'title', 0.59, 4.05, 7.6, 0.88, { fontFace: headingFont, fontSize: 24, color: COLORS.muted, lineSpacingMultiple: 1.12 }),
			placeholder('subtitle', 'body', 0.59, 4.98, 7.4, 0.3, { fontSize: 11, color: COLORS.muted }),
			placeholder('description', 'body', 0.59, 5.48, 7.1, 0.9, { fontSize: 11, lineSpacingMultiple: 1.35 }),
		],
	});
	const branding = [
		image('logo-yellow', 10.54, 0.45, 1, 0.5),
		{ text: { text: 'SABIC |', options: { ...text, x: 10.48, y: 7.35, w: 0.65, h: 0.2, fontSize: 8, align: 'right' } } },
	];
	const headings = [
		placeholder('title', 'title', 0.59, 0.99, 10.6, 0.52, { fontFace: headingFont, fontSize: 18 }),
		placeholder('subtitle', 'body', 0.59, 1.52, 10.6, 0.25, { fontSize: 8 }),
	];
	const column = (name, x, w, fontSize = 12) => [
		placeholder(`${name}Heading`, 'body', x, 1.96, w, 0.52, { fontSize: 12, color: COLORS.cyan, bold: true }),
		placeholder(name, 'body', x, 2.57, w, 4.2, { fontSize, paraSpaceAfterPt: 12, lineSpacingMultiple: 1.25 }),
	];
	const layouts = {
		SABIC_CONTENT: column('body', 0.59, 10.6, 16),
		SABIC_TWO_COLUMNS: [...column('left', 0.59, 5.08), ...column('right', 6.16, 5.08)],
		SABIC_THREE_COLUMNS: [...column('left', 0.59, 3.27, 10), ...column('middle', 4.32, 3.27, 10), ...column('right', 8.05, 3.19, 10)],
		SABIC_FINANCIAL_TABLE: [],
		SABIC_TEXT_AND_TABLE: column('left', 0.59, 5.08, 11),
	};
	for (const [title, objects] of Object.entries(layouts)) {
		pptx.defineSlideMaster({
			title, background: { color: 'FFFFFF' }, objects: [...branding, ...headings, ...objects],
			slideNumber: { x: 11.16, y: 7.35, w: 0.25, h: 0.2, fontFace: bodyFont, fontSize: 8, color: COLORS.text, margin: 0 },
		});
	}
	return pptx;
}

function addFinancialTable(slide, rows, { x = 0.59, y = 2.1, w = 10.6, colW = [6.4, 2.1, 2.1], fontSize = 11, fontFace = 'Helvetica Neue' } = {}) {
	if (!Array.isArray(rows) || rows.length < 2 || rows.some(row => row.length !== colW.length)) {
		throw new Error('Financial table needs a header and body rows with one cell per column width.');
	}
	const blank = { type: 'none' };
	const cells = rows.map((row, r) => row.map((value, c) => ({
		text: String(value),
		options: {
			align: c === 0 ? 'left' : 'right',
			color: c === 1 ? COLORS.cyan : COLORS.text,
			bold: r === 0 || r === rows.length - 1 || c === 1,
			border: [blank, blank, { type: 'solid', color: r === 0 ? COLORS.cyan : COLORS.rule, pt: r === 0 || r === rows.length - 1 ? 1.3 : 0.5 }, blank],
		},
	})));
	slide.addTable(cells, { x, y, w, colW, fontSize, fontFace, rowH: 0.34, margin: [5, 6, 5, 6], color: COLORS.text, fill: 'FFFFFF', autoPage: false });
}

module.exports = { createSabicPresentation, addFinancialTable, COLORS, SIZE };
