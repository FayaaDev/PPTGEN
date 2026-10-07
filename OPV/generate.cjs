const fs = require('node:fs/promises');
const path = require('node:path');
const PptxGenJS = require('../dist/pptxgen.cjs.js');
const JSZip = require('jszip');

const C = { ink: '0C2A33', teal: '136F7A', blue: '2B6F94', green: '2D7A62', body: '3A4F58', muted: '62767E', mint: 'D2E7E8', ice: 'DBE7F2', paper: 'F2F6F7', rule: 'D4DEE1', white: 'FFFFFF' };
const FONT = 'Franklin Gothic Book';
const WIDTH = 13.333333, M = 0.45, CW = WIDTH - 2 * M, GAP = 0.2;
const HALF = (CW - GAP) / 2, THIRD = (CW - 2 * GAP) / 3;
const names = ['Cover', 'Section', 'Content', 'Comparison', 'Three cards', 'Table', 'Chart', 'Image', 'Timeline', 'Summary', 'Appendix', 'References'];
const layouts = names.map((name, i) => `OPV ${String(i + 1).padStart(2, '0')} - ${name}`);

function ph(name, x, y, w, h, fontSize = 20, extra = {}) {
	return { placeholder: { text: '', options: { name, type: name === 'title' ? 'title' : 'body', x, y, w, h, fontFace: FONT, fontSize, color: C.body, margin: 0, breakLine: false, valign: 'top', paraSpaceAfterPt: 8, ...extra } } };
}
function rect(x, y, w, h, color) {
	return { rect: { x, y, w, h, fill: { color }, line: { color, transparency: 100 } } };
}
function panel(x, y, w, h, color = C.paper) {
	return [rect(x, y, w, h, color), rect(x, y, 0.035, h, C.teal)];
}
function shared(dark = false) {
	return [
		rect(M, 6.98, CW, 0.008, dark ? '31535E' : C.rule),
		ph('footer', M, 7.1, CW - 0.6, 0.2, 10, { color: dark ? 'A8C6CC' : C.muted }),
		...(!dark ? [ph('source', M, 6.55, CW, 0.31, 10, { color: C.muted, paraSpaceAfterPt: 0 })] : []),
	];
}
function heading() {
	return [ph('section', M, 0.32, CW, 0.19, 10, { color: C.teal, charSpacing: 1.3 }),
		rect(M, 0.64, 0.6, 0.035, C.teal),
		ph('title', M, 0.79, CW, 0.72, 30, { bold: true, color: C.ink }),
		ph('subtitle', M, 1.55, CW, 0.31, 14, { color: C.muted })];
}
function takeaway(y = 5.79) {
	return [...panel(M, y, CW, 0.58, C.mint), ph('takeaway', M + 0.17, y + 0.12, CW - 0.34, 0.34, 17, { color: C.ink, bold: true })];
}

async function createOpvPresentation() {
	const source = await JSZip.loadAsync(await fs.readFile(path.join(__dirname, 'Polio_2026_7.pptx')));
	const image = async file => {
		if (!source.file(file)) throw new Error(`Original artwork missing: ${file}`);
		return `image/png;base64,${await source.file(file).async('base64')}`;
	};
	const ghost = await image('ppt/media/image1.png');
	const capsid = await image('ppt/media/image11.png');
	const pptx = new PptxGenJS();
	pptx.layout = 'LAYOUT_WIDE';
	pptx.author = 'VPD Department, Public Health Authority';
	pptx.subject = 'OPV compact teaching master — reusable layouts';
	pptx.title = 'OPV | Compact Teaching Master';
	pptx.company = 'Public Health Authority';
	pptx.lang = 'en-GB';
	pptx.theme = { headFontFace: FONT, bodyFontFace: FONT, lang: 'en-GB' };
	const define = (index, objects, dark = false) => pptx.defineSlideMaster({
		title: layouts[index], background: { color: dark ? C.ink : C.white },
		objects: [...shared(dark), ...objects],
		slideNumber: { x: WIDTH - M - 0.45, y: 7.09, w: 0.45, h: 0.21, fontFace: FONT, fontSize: 10, color: dark ? 'A8C6CC' : C.muted, align: 'right', margin: 0 },
	});
	for (const index of [0, 1]) define(index, [
		{ image: { data: ghost, x: 9.05, y: 1.88, w: 3.7, h: 3.692, altText: 'Poliovirus capsid and RNA illustration from the original OPV deck' } },
		ph('section', 0.6, 0.65, 8, 0.24, 11, { color: 'A8C6CC', charSpacing: 1.8 }),
		rect(0.6, 1.91, 1.05, 0.045, C.teal),
		ph('title', 0.6, 2.18, 8.1, 1.4, index === 0 ? 42 : 38, { color: C.white, bold: true }),
		ph('subtitle', 0.6, 3.76, 7.9, 0.85, 22, { color: 'BED2D7' }),
		ph('presenter', 0.6, 5.1, 8, 0.37, 20, { color: C.white, bold: true }),
		ph('affiliation', 0.6, 5.58, 8, 0.65, 16, { color: 'BED2D7' }),
	], true);
	define(2, [...heading(), ph('body', M, 2.13, CW - 0.3, 3.35, 22), ...takeaway()]);
	define(3, [...heading(), ...[0, 1].flatMap(i => {
		const x = M + i * (HALF + GAP), name = i === 0 ? 'left' : 'right';
		return [...panel(x, 2.08, HALF, 3.48, i ? C.mint : C.paper),
			ph(`${name}Heading`, x + 0.2, 2.29, HALF - 0.4, 0.45, 23, { bold: true, color: i ? C.teal : C.ink }),
			ph(name, x + 0.2, 2.97, HALF - 0.4, 2.31, 20)];
	}), ...takeaway()]);
	define(4, [...heading(), ...[0, 1, 2].flatMap(i => {
		const x = M + i * (THIRD + GAP), name = ['left', 'middle', 'right'][i];
		return [...panel(x, 2.08, THIRD, 3.48, [C.mint, C.paper, C.ice][i]),
			ph(`${name}Heading`, x + 0.18, 2.28, THIRD - 0.36, 0.68, 22, { bold: true, color: [C.teal, C.ink, C.blue][i] }),
			ph(name, x + 0.18, 3.11, THIRD - 0.36, 2.17, 20)];
	}), ...takeaway()]);
	define(5, [...heading(), ph('table', M, 2.06, CW, 3.49, 16, { type: 'table' }), ...takeaway()]);
	for (const index of [6, 7]) define(index, [...heading(),
		ph('visual', M, 2.08, 7.88, 3.48, 16, { type: index === 6 ? 'chart' : 'image' }),
		...panel(8.55, 2.08, WIDTH - M - 8.55, 3.48, C.paper),
		ph('commentHeading', 8.75, 2.3, 3.68, 0.6, 22, { bold: true, color: C.teal }),
		ph('comment', 8.75, 3.08, 3.68, 2.23, 20), ...takeaway()]);
	define(8, [...heading(), ...[0, 1, 2, 3, 4].flatMap(i => {
		const y = 2.06 + i * 0.7;
		return [rect(M, y, 1.45, 0.57, i % 2 ? C.ice : C.mint),
			ph(`step${i + 1}`, M + 0.13, y + 0.13, 1.19, 0.3, 18, { color: C.teal, bold: true }),
			ph(`action${i + 1}`, 2.1, y + 0.11, 10.45, 0.43, 19)];
	}), ...takeaway()]);
	define(9, [...heading(), ...Array.from({ length: 6 }, (_, i) => {
		const x = M + (i % 2) * (HALF + GAP), y = 2.08 + Math.floor(i / 2) * 1.42;
		return [...panel(x, y, HALF, 1.24, i % 2 ? C.paper : C.mint),
			ph(`heading${i + 1}`, x + 0.17, y + 0.14, HALF - 0.34, 0.36, 20, { bold: true, color: C.ink }),
			ph(`body${i + 1}`, x + 0.17, y + 0.61, HALF - 0.34, 0.46, 17)];
	})]);
	define(10, [...heading(), rect(0, 0, 0.055, 7.5, C.green),
		ph('body', M, 2.1, CW, 3.36, 20), ...takeaway()]);
	define(11, [...heading(), ...[0, 1].flatMap(i => {
		const x = M + i * (HALF + GAP), name = i === 0 ? 'left' : 'right';
		return [ph(`${name}Heading`, x, 2.08, HALF, 0.37, 21, { bold: true, color: C.teal }),
			ph(name, x, 2.71, HALF - 0.1, 3.48, 18)];
	})]);
	return { pptx, capsid };
}

function fill(slide, name, text, options = {}) {
	slide.addText(text, { placeholder: name, ...options });
}
function bullets(slide, name, items) {
	fill(slide, name, items.map(text => ({ text, options: { bullet: { indent: 16 }, hanging: 4, breakLine: true } })), { paraSpaceAfterPt: 13 });
}

async function main() {
	const { pptx, capsid } = await createOpvPresentation();
	const add = (index, title, subtitle, section, source = '') => {
		const slide = pptx.addSlide({ masterName: layouts[index] });
		fill(slide, 'title', title);
		fill(slide, 'subtitle', subtitle);
		fill(slide, 'section', section);
		fill(slide, 'footer', 'VPD Department · Public Health Authority');
		if (index > 1) fill(slide, 'source', source);
		slide.addNotes(`Layout: ${layouts[index]}. Use Home > New Slide to choose this layout, or duplicate this example to retain its editable objects. Replace title, section, subtitle, footer and source. Keep content above the citation zone. Example text demonstrates design; check current guidance before use.`);
		return slide;
	};
	let s = add(0, 'Poliomyelitis', 'Preventive medicine residents\nteaching session', 'VACCINE PREVENTABLE DISEASES');
	fill(s, 'presenter', 'Abdullah Alshehri, MBBS, SBPM, MHA');
	fill(s, 'affiliation', 'Vaccine Preventable Diseases Department\nCommunicable Diseases Sector');

	s = add(1, 'Surveillance', 'Finding transmission before it becomes visible.', 'SECTION 04 / SURVEILLANCE');
	fill(s, 'presenter', 'AFP · Environmental · Immunodeficiency-associated');
	fill(s, 'affiliation', 'A coordinated approach to detecting poliovirus.');

	s = add(2, 'Learning objectives', 'Use a short action-led list for the teaching session.', 'SESSION OVERVIEW');
	bullets(s, 'body', ['Distinguish disease control, elimination and eradication.', 'Recognise the clinical spectrum and investigate acute flaccid paralysis.', 'Compare the roles of IPV and OPV.', 'Explain vaccine-derived poliovirus classification.', 'Apply surveillance standards and interpret their indicators.']);
	fill(s, 'takeaway', 'Teaching point · Lead with the learning goal, then the evidence.');

	s = add(3, 'Two vaccines, two different jobs', 'Use parallel headings and matched content for direct comparison.', 'PREVENTION / VACCINES', 'Source: CDC Pink Book, chapter 18. Example text adapted from the supplied OPV deck.');
	fill(s, 'leftHeading', 'IPV / Inactivated, injected');
	fill(s, 'rightHeading', 'OPV / Live attenuated, oral');
	bullets(s, 'left', ['Produces protective serum antibodies.', 'Protects against paralytic disease.', 'Contains no live poliovirus.']);
	bullets(s, 'right', ['Produces intestinal immunity.', 'Reduces shedding and transmission.', 'Requires attention to vaccine-virus risks.']);
	fill(s, 'takeaway', 'Teaching point · The sequential schedule uses both vaccines deliberately.');

	s = add(4, 'Prevention rests on three pillars', 'One heading, three coordinated areas of action.', 'PREVENTION / PROGRAMME', 'Source: supplied OPV presentation, prevention section.');
	for (const [name, title, items] of [
		['left', 'Immunisation', ['Maintain high coverage.', 'Deliver the schedule.', 'Close immunity gaps.']],
		['middle', 'Water, sanitation\nand hygiene', ['Reduce faecal–oral spread.', 'Protect water supplies.', 'Improve personal hygiene.']],
		['right', 'Surveillance\nand response', ['Investigate AFP promptly.', 'Monitor environmental samples.', 'Respond to detections.']],
	]) { fill(s, `${name}Heading`, title); bullets(s, name, items); }
	fill(s, 'takeaway', 'Teaching point · Vaccination, surveillance and response work together.');

	s = add(5, 'IPV and OPV / At a glance', 'Native PowerPoint table · edit cells directly.', 'PREVENTION / COMPARISON', 'Source: supplied OPV presentation; CDC Pink Book, chapter 18. Simplified teaching example.');
	fill(s, 'table', '');
	const rows = [['Feature', 'IPV', 'OPV'], ['Vaccine type', 'Inactivated', 'Live attenuated'], ['Route', 'Injection', 'Oral'], ['Main role', 'Protects against paralysis', 'Interrupts transmission'], ['Intestinal immunity', 'Limited when used alone', 'Stronger mucosal response'], ['Live vaccine virus', 'Absent', 'Replicates and is shed']];
	s.addTable(rows.map((row, i) => row.map(text => ({ text, options: i === 0 ? { fill: C.ink, color: C.white, bold: true } : { fill: i % 2 ? C.paper : C.white, color: C.body } }))), {
		x: M, y: 2.06, w: CW, h: 3.48, colW: [3.1, 4.6666665, 4.6666665], rowH: 0.58,
		fontFace: FONT, fontSize: 18, margin: [8, 12, 8, 12], border: { color: C.rule, pt: 0.6 }, valign: 'mid', autoPage: false,
	});
	fill(s, 'takeaway', 'Teaching point · Compare like-for-like features; emphasise the clinically useful distinction.');

	s = add(6, 'Coverage / Illustrative trend', 'Native PowerPoint chart · sample data, not surveillance figures.', 'DATA / TREND', 'Source: illustrative data created for this template. Replace values, units and reporting dates.');
	s.addChart(pptx.ChartType.bar, [{ name: 'Coverage', labels: ['Round 1', 'Round 2', 'Round 3', 'Round 4'], values: [72, 81, 88, 94] }], {
		placeholder: 'visual', x: M, y: 2.08, w: 7.88, h: 3.48,
		catAxisLabelFontFace: FONT, catAxisLabelFontSize: 14, valAxisLabelFontFace: FONT, valAxisLabelFontSize: 12,
		valAxisMinVal: 0, valAxisMaxVal: 100, valAxisMajorUnit: 20, valAxisLabelFormatCode: '0"%"',
		chartColors: [C.teal], showLegend: false, showTitle: false, showValue: true,
		dataLabelFormatCode: '0"%"', dataLabelColor: C.ink, dataLabelPosition: 'outEnd', dataLabelFormatCodeSourceLinked: false,
		dataLabelBkgrdColor: C.white, dataLabelFontFace: FONT, dataLabelFontSize: 14,
		showBorder: false, showCatName: false, showCatName: false, catAxisLineShow: false, valAxisLineShow: false,
		valGridLine: { color: C.rule, width: 0.5 }, showCatName: false, showSerName: false,
		showMarker: false, gapSizePct: 90,
	});
	fill(s, 'commentHeading', 'Read the trend');
	fill(s, 'comment', 'Compare the same measure across time.\n\nState the denominator, reporting period and data source.\n\nExplain what changed and why it matters.');
	fill(s, 'takeaway', 'Teaching point · Put interpretation beside the chart, not inside the plot.');

	s = add(7, 'Poliovirus / Structure and serotypes', 'Image-led explanation with a compact commentary column.', 'THE VIRUS / STRUCTURE', 'Source: original capsid illustration and teaching content from the supplied OPV deck.');
	s.addImage({ data: capsid, placeholder: 'visual', x: 2.51, y: 2.08, w: 3.48, h: 3.4727, altText: 'Poliovirus cutaway showing a geometric protein capsid surrounding RNA' });
	fill(s, 'commentHeading', 'What to look for');
	fill(s, 'comment', 'A protein capsid encloses a single-stranded RNA genome.\n\nThere is no lipid envelope.\n\nImmunity is serotype-specific.');
	fill(s, 'takeaway', 'Teaching point · Use the image to explain one concept at a time.');

	s = add(8, 'AFP investigation / A time-critical process', 'Five steps on a consistent timeline.', 'SURVEILLANCE / WORKFLOW', 'Source: supplied OPV presentation, AFP investigation timeline. Verify against current national guidance.');
	for (const [i, when, action] of [
		[1, 'Immediately', 'Notify the suspected AFP case through the designated route.'],
		[2, '≤ 48 hours', 'Investigate the case and begin arranging specimen collection.'],
		[3, '≤ 14 days', 'Collect two stool specimens, at least 24 hours apart.'],
		[4, 'Day 60–90', 'Follow up and assess for residual paralysis.'],
		[5, '≤ 90 days', 'Complete expert review and final case classification.'],
	]) { fill(s, `step${i}`, when); fill(s, `action${i}`, action); }
	fill(s, 'takeaway', 'Teaching point · The clock starts at onset of paralysis.');

	s = add(9, 'What to take away', 'Six concise messages to close the teaching session.', 'SESSION SUMMARY', 'Source: supplied OPV presentation. Shortened example copy for the reusable summary layout.');
	for (const [i, title, body] of [
		[1, 'Most infection is invisible', 'Cases alone cannot measure transmission.'],
		[2, 'Two vaccines, different roles', 'IPV protects; OPV adds intestinal immunity.'],
		[3, 'Coverage determines risk', 'Missed children leave gaps in population immunity.'],
		[4, 'Classification needs context', 'Genetics and epidemiology answer different questions.'],
		[5, 'Surveillance quality matters', 'Timely investigation and adequate specimens are essential.'],
		[6, 'The last mile is operational', 'Access, trust and sustained response remain critical.'],
	]) { fill(s, `heading${i}`, title); fill(s, `body${i}`, body); }

	s = add(10, 'Appendix / Definitions and supporting notes', 'Use the appendix for detail that supports the core teaching narrative.', 'APPENDIX / SUPPORTING DETAIL', 'Source: replace with the relevant guideline, edition, publication date and page.');
	bullets(s, 'body', ['Define technical terms precisely and cite their source.', 'Record which guidance or case definition is being used.', 'Keep contextual explanations separate from the headline conclusion.', 'Use the main deck for the key message and the appendix for detail.']);
	fill(s, 'takeaway', 'Reference note · Add edition and date when definitions differ between sources.');

	s = add(11, 'References / Source register', 'Keep source names, editions and dates easy to scan.', 'APPENDIX / REFERENCES', 'Example source register adapted from the supplied OPV deck. Verify editions and links before presenting.');
	fill(s, 'leftHeading', 'Clinical and vaccine guidance');
	fill(s, 'left', '01  CDC Pink Book\nChapter 18: Poliomyelitis · reviewed May 2024\n\n02  WHO position paper\nPolio vaccines · June 2022\n\n03  Jafari H et al.\nEfficacy of IPV in India · Science 2014;345:922–925');
	fill(s, 'rightHeading', 'Surveillance and national guidance');
	fill(s, 'right', '04  GPEI\nGlobal AFP surveillance guidance\n\n05  Global Polio Laboratory Network\nGuidance Paper 5 · version 2.0 · March 2024\n\n06  Public Health Authority\nNational schedule and notification guidance');

	const pptxPath = path.join(__dirname, 'OPV_Compact_Examples.pptx');
	await pptx.writeFile({ fileName: pptxPath });
	const zip = await JSZip.loadAsync(await fs.readFile(pptxPath));
	// Match the Office theme accents to the design, including newly inserted charts/shapes.
	for (const filename of Object.keys(zip.files).filter(name => /^ppt\/theme\/theme\d+\.xml$/.test(name))) {
		let xml = await zip.file(filename).async('string');
		for (const [tag, color] of Object.entries({ dk1: C.ink, lt1: C.white, dk2: C.body, lt2: C.paper, accent1: C.teal, accent2: C.blue, accent3: C.green, accent4: C.mint, accent5: C.ice, accent6: C.muted, hlink: C.blue, folHlink: C.teal })) {
			xml = xml.replace(new RegExp(`<a:${tag}>[\\s\\S]*?</a:${tag}>`), `<a:${tag}><a:srgbClr val="${color}"/></a:${tag}>`);
		}
		zip.file(filename, xml);
	}
	await fs.writeFile(pptxPath, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
	const types = await zip.file('[Content_Types].xml').async('string');
	const presentationType = 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml';
	if (!types.includes(presentationType)) throw new Error('PPTX main content type missing');
	zip.file('[Content_Types].xml', types.replace(presentationType, 'application/vnd.openxmlformats-officedocument.presentationml.template.main+xml'));
	await fs.writeFile(path.join(__dirname, 'OPV_Compact_Master.potx'), await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
	console.log('Created OPV/OPV_Compact_Examples.pptx and OPV/OPV_Compact_Master.potx (12 example slides, 12 named layouts).');
}

module.exports = { createOpvPresentation, layouts };
if (require.main === module) main().catch(error => { console.error(error); process.exitCode = 1; });
