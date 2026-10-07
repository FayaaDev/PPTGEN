const path = require('node:path');
const fs = require('node:fs/promises');
const JSZip = require('jszip');
const { createSabicPresentation, addFinancialTable } = require('./template.cjs');

async function main() {
	const bodyFont = process.env.SABIC_BODY_FONT || 'Helvetica Neue';
	const pptx = createSabicPresentation({ bodyFont, headingFont: process.env.SABIC_HEADING_FONT || bodyFont });
	const add = (layout, title, subtitle = '') => {
		const slide = pptx.addSlide({ masterName: layout });
		slide.addText(title, { placeholder: 'title' });
		slide.addText(subtitle, { placeholder: 'subtitle' });
		return slide;
	};
	let slide = add('SABIC_COVER', 'SAUDI BASIC INDUSTRIES\nCORPORATION (SABIC)', '(A SAUDI JOINT STOCK COMPANY)');
	slide.addText('PRESENTATION TITLE\nREPORTING PERIOD / DATE\nDEPARTMENT OR PRESENTER', { placeholder: 'description' });
	slide.addNotes('Replace the title, reporting period and presenter. Cover artwork is extracted from the supplied PDF; text remains editable.');

	slide = add('SABIC_FINANCIAL_TABLE', 'CONTENTS', 'Presentation overview');
	addFinancialTable(slide, [
		['Section', 'Slide'], ['Executive summary', '4'], ['Business overview', '5'],
		['Key priorities', '6'], ['Financial performance', '7'], ['Notes and supporting information', '8'],
	], { colW: [9.1, 1.5], fontSize: 15, fontFace: bodyFont });
	slide.addNotes('Edit section names and slide numbers to match your presentation. This is a native PowerPoint table.');

	slide = add('SABIC_COVER', 'SECTION TITLE', 'SECTION 01');
	slide.addText('Short description of this section\nReplace with your own content', { placeholder: 'description' });

	slide = add('SABIC_CONTENT', 'EXECUTIVE SUMMARY', 'Replace sample copy with your key messages');
	slide.addText('KEY MESSAGES', { placeholder: 'bodyHeading' });
	slide.addText('01  Lead with the most important conclusion.\n\n02  Explain the business context and supporting evidence.\n\n03  State the next decision, action or milestone.', { placeholder: 'body' });

	slide = add('SABIC_TWO_COLUMNS', 'BUSINESS OVERVIEW', 'Two-column layout based on the financial report');
	slide.addText('1.  BUSINESS CONTEXT', { placeholder: 'leftHeading' });
	slide.addText('Introduce the business, reporting period and scope of the presentation.\n\nSummarise the current situation with a short paragraph and supporting facts.\n\nReplace this sample copy with your own content.', { placeholder: 'left' });
	slide.addText('2.  PRIORITIES AND ACTIONS', { placeholder: 'rightHeading' });
	slide.addText('Describe the priorities for the next reporting period.\n\nExplain the actions, owners and expected milestones.\n\nUse the same headings, margins and spacing throughout the deck.', { placeholder: 'right' });

	slide = add('SABIC_THREE_COLUMNS', 'KEY PRIORITIES', 'Three-column layout based on the notes pages');
	for (const [name, heading, body] of [
		['left', '1.  PERFORMANCE', 'Describe the objective and the measures used to track performance.\n\nAdd a concise explanation of the current position and the intended outcome.'],
		['middle', '2.  OPERATIONS', 'Describe the operational priority and the action required.\n\nAdd the responsible owner and the target milestone.'],
		['right', '3.  OUTLOOK', 'Describe the opportunity, dependency or risk that needs attention.\n\nState the next step and when progress will be reviewed.'],
	]) {
		slide.addText(heading, { placeholder: `${name}Heading` });
		slide.addText(body, { placeholder: name });
	}

	slide = add('SABIC_FINANCIAL_TABLE', 'FINANCIAL PERFORMANCE', 'Illustrative placeholders — replace with your figures and units');
	addFinancialTable(slide, [
		['Metric', 'Current period', 'Prior period'], ['Revenue', '—', '—'], ['Cost of sales', '—', '—'],
		['Gross profit', '—', '—'], ['Operating expenses', '—', '—'], ['Income from operations', '—', '—'],
		['Finance income / (costs)', '—', '—'], ['Zakat and income tax', '—', '—'], ['Net income / (loss)', '—', '—'],
	], { fontFace: bodyFont });
	slide.addText('Source: [insert source]. All amounts in [insert units].', { x: 0.59, y: 6.75, w: 10.6, h: 0.25, fontFace: bodyFont, fontSize: 8, color: '575756', margin: 0 });
	slide.addNotes('Native editable table. Current-period figures use SABIC cyan; prior-period figures use grey. No financial figures are invented.');

	slide = add('SABIC_TEXT_AND_TABLE', 'NOTES AND SUPPORTING INFORMATION', 'Text and table layout based on the disclosure pages');
	slide.addText('1.  EXPLANATORY NOTE', { placeholder: 'leftHeading' });
	slide.addText('Describe the scope, basis of preparation or relevant context for the table.\n\nExplain any assumptions, definitions or changes in presentation.\n\nReplace the table headings and figures with the information relevant to your presentation.', { placeholder: 'left' });
	addFinancialTable(slide, [
		['Category', 'Current', 'Prior'], ['Category A', '—', '—'], ['Category B', '—', '—'],
		['Category C', '—', '—'], ['Total', '—', '—'],
	], { x: 6.16, y: 2.1, w: 5.08, colW: [2.88, 1.1, 1.1], fontSize: 10, fontFace: bodyFont });
	slide.addNotes('Use Home > New Slide to choose the included SABIC layouts, or duplicate a sample slide. All content is editable; branding is held in the slide layouts.');

	const fileName = path.join(__dirname, 'SABIC-Template.pptx');
	await pptx.writeFile({ fileName });
	const zip = await JSZip.loadAsync(await fs.readFile(fileName));
	const types = await zip.file('[Content_Types].xml').async('string');
	zip.file('[Content_Types].xml', types.replace(
		'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',
		'application/vnd.openxmlformats-officedocument.presentationml.template.main+xml',
	));
	await fs.writeFile(path.join(__dirname, 'SABIC-Template.potx'), await zip.generateAsync({ type: 'nodebuffer' }));
	console.log('Created SABIC/SABIC-Template.pptx and SABIC/SABIC-Template.potx (8 slides).');
}

main().catch(error => { console.error(error); process.exitCode = 1; });
