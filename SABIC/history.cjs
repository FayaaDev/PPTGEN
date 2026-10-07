const fs = require('node:fs/promises');
const path = require('node:path');
const assert = require('node:assert/strict');
const { createSabicPresentation, COLORS } = require('./template.cjs');

const SOURCES = [
	['About SABIC', 'https://www.sabic.com/en/about'],
	['Corporate Profile', 'https://www.sabic.com/en/about/corporate-profile'],
	['50 Years: anniversary timeline', 'https://www.sabic.com/en/about/50-years'],
	['2020 Annual Report: Environmental Capital', 'https://www.sabic.com/en/reports/annual-2020/performance/environmental-capital'],
	['Aramco acquisition completion announcement', 'https://www.sabic.com/en/news/23759-sabic-announces-the-completion-of-saudi-aramco-acquisition-of-public-investment-fund'],
];

async function main() {
	const assetDir = process.argv[2];
	assert(assetDir, 'Usage: node SABIC/history.cjs <directory containing sabic-history-1977.png and sabic-history-1983.png>');
	const photos = await Promise.all([1977, 1983].map(async year =>
		`image/png;base64,${(await fs.readFile(path.join(assetDir, `sabic-history-${year}.png`))).toString('base64')}`));
	const pptx = createSabicPresentation();
	pptx.title = 'History of SABIC: 1976–2026';
	pptx.subject = 'History of SABIC, based on its official website';
	pptx.author = 'Presentation prepared from SABIC public sources';
	const fontFace = 'Helvetica Neue';
	const text = (slide, content, x, y, w, h, options = {}) => slide.addText(content, {
		x, y, w, h, fontFace, fontSize: 16, margin: 0, color: COLORS.text, align: 'left', valign: 'top', ...options,
	});
	const line = (slide, x, y, w, color = COLORS.cyan) => slide.addShape(pptx.ShapeType.line, { x, y, w, h: 0, line: { color, width: 1 } });
	function add(title, subtitle, sources, notes = '') {
		const slide = pptx.addSlide({ masterName: 'SABIC_FINANCIAL_TABLE' });
		slide.addText(title, { placeholder: 'title' });
		slide.addText(subtitle, { placeholder: 'subtitle', fontSize: 10 });
		text(slide, `Sources: ${sources.map(n => `[${n}] ${SOURCES[n - 1][0]}`).join(' · ')}`, 0.59, 7.01, 9.65, 0.22, { fontSize: 7.5, color: COLORS.muted });
		slide.addNotes(`${notes}\n\nOfficial sources (accessed 7 October 2026):\n${sources.map(n => `[${n}] ${SOURCES[n - 1][0]}\n${SOURCES[n - 1][1]}`).join('\n\n')}`);
		return slide;
	}
	function card(slide, date, heading, body, x, y, w = 3.23) {
		text(slide, date, x, y, w, 0.55, { fontSize: 30, color: COLORS.cyan });
		line(slide, x, y + 0.7, w);
		text(slide, heading, x, y + 0.89, w, 0.65, { fontSize: 15, bold: true });
		text(slide, body, x, y + 1.7, w, 1.3, { fontSize: 14, lineSpacingMultiple: 1.15 });
	}

	let slide = pptx.addSlide({ masterName: 'SABIC_COVER' });
	slide.addText('HISTORY OF SABIC\n1976–2026', { placeholder: 'title' });
	slide.addText('FIFTY YEARS OF INDUSTRIAL DEVELOPMENT', { placeholder: 'subtitle' });
	slide.addText('From a Saudi industrial vision to a global chemicals company\nBased on SABIC’s official website', { placeholder: 'description' });
	slide.addNotes(`Overview of selected milestones in SABIC's history. Research accessed 7 October 2026.\n${SOURCES.map(([name, url], i) => `[${i + 1}] ${name}: ${url}`).join('\n')}`);

	slide = add('FIFTY YEARS AT A GLANCE', 'Selected milestones: foundation, production, expansion and transformation', [2, 3, 4, 5]);
	for (const [i, date, heading, body] of [
		[0, '1976', 'FOUNDATION', 'Established by royal decree in Saudi Arabia.'],
		[1, '1983–1985', 'INDUSTRIAL START-UP', 'Joint-venture production and exports begin; more plants follow.'],
		[2, '2002–2007', 'GLOBAL EXPANSION', 'DSM, Huntsman UK and GE Plastics acquisitions expand capabilities.'],
		[3, '2019', 'CIRCULAR MATERIALS', 'TRUCIRCLE™ launches at the end of the year.'],
		[4, '2020', 'ARAMCO PARTNERSHIP', 'Saudi Aramco acquires a 70% stake.'],
		[5, '2026', 'GOLDEN JUBILEE', 'SABIC marks its 50th anniversary.'],
	]) {
		const x = 0.59 + (i % 3) * 3.73;
		const y = 2.1 + Math.floor(i / 3) * 2.35;
		text(slide, date, x, y, 3.14, 0.5, { fontSize: 25, color: COLORS.cyan });
		line(slide, x, y + 0.61, 3.14);
		text(slide, heading, x, y + 0.79, 3.14, 0.32, { fontSize: 12, bold: true });
		text(slide, body, x, y + 1.25, 3.14, 0.84, { fontSize: 13 });
	}

	slide = add('1976: AN INDUSTRIAL VISION', 'Founded by royal decree in September 1976', [2, 3], 'The royal decree was signed by King Khalid bin Abdulaziz Al-Saud. The original purpose included using oil-extraction by-products to produce chemicals, polymers and fertilizers for export and local downstream industries. Photo: SABIC’s first Riyadh headquarters, 1977, from its 50 Years page.');
	text(slide, 'TURNING RESOURCES\nINTO VALUE', 0.59, 2.15, 4.6, 1, { fontSize: 26, color: COLORS.cyan });
	text(slide, 'SABIC was created to convert oil-extraction by-products into value-added chemicals, polymers and fertilizers.', 0.59, 3.42, 4.6, 1.3);
	text(slide, 'The ambition: develop Saudi downstream industries and serve export markets.', 0.59, 5.03, 4.6, 1.1);
	slide.addImage({ data: photos[0], x: 5.62, y: 2.2, w: 5.58, h: 2.986 });
	text(slide, 'SABIC’s first headquarters in Riyadh, 1977', 5.62, 5.36, 5.58, 0.4, { fontSize: 10, color: COLORS.muted });

	slide = add('1983–1985: FROM PLANS TO PRODUCTION', 'Joint ventures turned the industrial vision into operating businesses', [2, 3], 'The Royal Commission for Jubail and Yanbu built infrastructure while SABIC built factories and acquired technology through joint ventures. Joint-venture production began in 1983 with Hadeed, Al-Bayroni and Ar-Razi; Gas followed in 1984. Sadaf, Yanpet, Petrokemya, Kemya, Sharq and Ibn Sina began production in 1985. SABIC’s anniversary timeline dates its first exports to 1983. Photo: first exports, from the anniversary page.');
	slide.addImage({ data: photos[1], x: 0.59, y: 2.2, w: 5.45, h: 2.917 });
	text(slide, 'First exports outside the Kingdom, 1983', 0.59, 5.28, 5.45, 0.35, { fontSize: 10, color: COLORS.muted });
	for (const [i, year, detail] of [
		[0, '1983', 'Production begins: Hadeed, Al-Bayroni and Ar-Razi.'],
		[1, '1984', 'Gas joins the production network.'],
		[2, '1985', 'Six more joint ventures begin production.'],
	]) {
		const y = 2.2 + i * 1.38;
		text(slide, year, 6.43, y, 1.2, 0.5, { fontSize: 24, color: COLORS.cyan });
		text(slide, detail, 7.91, y + 0.03, 3.3, 0.91, { fontSize: 14 });
		line(slide, 6.43, y + 1.1, 4.77, COLORS.rule);
	}

	slide = add('1984: OPENING TO PUBLIC INVESTORS', 'SABIC’s first initial public offering', [3], 'The 50 Years page dates SABIC’s first IPO to 1984 and describes it as expanding the shareholder base and strengthening the company’s financial-market presence. This slide describes that historical milestone, not a claim about IPO proceeds or shares offered.');
	text(slide, '1984', 0.59, 2.3, 4.2, 1.4, { fontSize: 76, color: COLORS.cyan });
	text(slide, 'FIRST IPO', 0.59, 4.05, 4.2, 0.6, { fontSize: 24, bold: true });
	line(slide, 5.14, 2.25, 6.05);
	text(slide, 'A broader shareholder base', 5.14, 2.7, 6.05, 0.6, { fontSize: 22 });
	text(slide, 'The offering expanded participation in SABIC’s industrial growth and strengthened its presence in the financial market.', 5.14, 3.64, 6.05, 1.45, { fontSize: 18 });
	text(slide, 'A new chapter in the company’s development.', 5.14, 5.56, 6.05, 0.62, { fontSize: 18, color: COLORS.cyan });

	slide = add('2002–2007: BUILDING A GLOBAL PORTFOLIO', 'Three acquisitions expanded manufacturing reach and materials capabilities', [2], 'DSM’s petrochemical business brought facilities in the Netherlands, Germany and the UK. Huntsman Petrochemicals (UK) was renamed SABIC UK Petrochemicals. GE Plastics expanded advanced materials and value-added products. These are historical acquisitions; they are not a statement that the acquired portfolio is unchanged today.');
	card(slide, '2002', 'DSM PETROCHEMICALS', 'European manufacturing facilities added a new platform for global expansion.', 0.59, 2.22);
	card(slide, '2006', 'HUNTSMAN UK', 'The UK petrochemicals acquisition added capacity to SABIC’s European operations.', 4.32, 2.22);
	card(slide, '2007', 'GE PLASTICS', 'The acquisition broadened advanced-materials and value-added product capabilities.', 8.05, 2.22, 3.15);
	text(slide, 'Broader geography. Greater capacity. More specialized materials.', 0.59, 6.11, 10.6, 0.52, { fontSize: 20, color: COLORS.cyan });

	slide = add('2019: A NEW CHAPTER IN CIRCULARITY', 'TRUCIRCLE™ launched at the end of 2019', [4], 'SABIC’s 2020 Annual Report states that TRUCIRCLE launched at the end of 2019. It described mechanical recycling, certified circular products, certified bio-based renewable products and design for recyclability. The diagram is a simplified illustration of the circular-materials ambition, not a claim that every material is fully circular.');
	text(slide, 'TRUCIRCLE™', 0.59, 2.14, 10.6, 0.74, { fontSize: 35, color: COLORS.cyan });
	text(slide, 'A portfolio of solutions intended to help close the loop on plastic waste.', 0.59, 3.12, 10.6, 0.72, { fontSize: 22 });
	for (const [i, heading, body] of [
		[0, 'RECOVER', 'Mechanical recycling and certified circular feedstocks.'],
		[1, 'REMAKE', 'Materials that incorporate recycled or renewable inputs.'],
		[2, 'REDESIGN', 'Design for recyclability and collaboration across the value chain.'],
	]) {
		const x = 0.59 + i * 3.73;
		line(slide, x, 4.34, 3.14);
		text(slide, heading, x, 4.65, 3.14, 0.48, { fontSize: 20, color: COLORS.cyan });
		text(slide, body, x, 5.36, 3.14, 1.04, { fontSize: 14 });
	}

	slide = add('2020: THE SAUDI ARAMCO PARTNERSHIP', 'Saudi Aramco acquired PIF’s 70% stake in SABIC in June 2020', [1, 5], 'The completion announcement states that the parties signed the final agreement on 16 June 2020 and that Saudi Aramco became a 70% shareholder. SABIC was positioned as Aramco’s chemicals growth platform. The current About page reports 70% Aramco ownership and 30% publicly traded. The chart depicts the share split, not market capitalization.');
	slide.addChart(pptx.ChartType.doughnut, [{ name: 'Share ownership', labels: ['Saudi Aramco', 'Publicly traded'], values: [70, 30] }], {
		x: 0.59, y: 2.13, w: 4.75, h: 4.25, holeSize: 70,
		chartColors: [COLORS.blue, COLORS.yellow], showLegend: false, showTitle: false,
		showValue: false, showPercent: false, showBorder: false,
	});
	text(slide, '70%', 1.64, 3.74, 2.64, 0.7, { fontSize: 40, color: COLORS.blue, align: 'center' });
	text(slide, 'Saudi Aramco', 1.64, 4.49, 2.64, 0.4, { fontSize: 13, align: 'center' });
	text(slide, '30% publicly traded', 0.59, 6.38, 4.75, 0.45, { fontSize: 14, align: 'center' });
	text(slide, 'A chemicals growth platform', 5.92, 2.39, 5.29, 0.72, { fontSize: 23, color: COLORS.cyan });
	text(slide, 'The acquisition connected SABIC’s chemicals capabilities with Saudi Aramco’s global energy business.', 5.92, 3.48, 5.29, 1.25, { fontSize: 18 });
	text(slide, 'SABIC remained a publicly traded company, with 30% of its shares traded on the Saudi stock exchange.', 5.92, 5.05, 5.29, 1.3, { fontSize: 17 });

	slide = add('GROWTH IN REPORTED PRODUCTION', 'Historical snapshots from SABIC’s Corporate Profile', [2], 'SABIC reports total production of 6.3 million metric tons in 1985 and 55.5 million metric tons in 2025. These are historical totals. Portfolio composition changed over time; the comparison should not be read as like-for-like organic growth.');
	slide.addChart(pptx.ChartType.bar, [{ name: 'Million metric tons', labels: ['1985', '2025'], values: [6.3, 55.5] }], {
		x: 0.78, y: 2.12, w: 7.35, h: 4.28, barDir: 'col',
		catAxisLabelFontFace: fontFace, catAxisLabelFontSize: 15,
		valAxisLabelFontFace: fontFace, valAxisLabelFontSize: 10, valAxisMinVal: 0, valAxisMaxVal: 60, valAxisMajorUnit: 20,
		chartColors: [COLORS.cyan], showLegend: false, showTitle: false, showValue: true,
		dataLabelPosition: 'outEnd', dataLabelFormatCode: '0.0', dataLabelColor: COLORS.text, dataLabelFontSize: 16,
		showCatName: false, showBorder: false,
	});
	text(slide, '55.5', 8.62, 2.58, 2.57, 0.9, { fontSize: 43, color: COLORS.cyan });
	text(slide, 'million metric tons\nreported in 2025', 8.62, 3.69, 2.57, 1.0, { fontSize: 16 });
	text(slide, 'Historical totals; portfolio composition changed over time.', 0.78, 6.55, 10.4, 0.3, { fontSize: 9, color: COLORS.muted });

	slide = add('2026: SABIC’S GOLDEN JUBILEE', 'A national industrial vision with a global footprint', [1, 3], 'SABIC’s anniversary page marks fifty years in 2026. The About page reports more than 26,000 employees and customers in more than 140 countries. The anniversary page reports 20 technology centers and service locations. These are website snapshots accessed 7 October 2026; they are not historical figures for 1976.');
	for (const [i, figure, label] of [
		[0, '50', 'YEARS\n1976–2026'],
		[1, '26,000+', 'EMPLOYEES\nAROUND THE WORLD'],
		[2, '140+', 'COUNTRIES\nSERVED'],
	]) {
		const x = 0.59 + i * 3.73;
		text(slide, figure, x, 2.47, 3.2, 1.02, { fontSize: i === 1 ? 42 : 58, color: COLORS.cyan });
		line(slide, x, 3.81, 3.2);
		text(slide, label, x, 4.13, 3.2, 0.9, { fontSize: 17, bold: true });
	}
	text(slide, 'SABIC’s principal headquarters remain in Riyadh, with a worldwide manufacturing, sales and innovation network.', 0.59, 5.76, 10.6, 0.88, { fontSize: 18 });

	slide = add('THE ENDURING THREAD IN SABIC’S HISTORY', 'From industrial foundations to global capabilities', [2, 3, 4, 5], 'This slide synthesizes the historical narrative: resource value creation in 1976; joint ventures, production and exports in the 1980s; international acquisitions from 2002; circular-materials initiatives in 2019; Aramco ownership in 2020; and the 50th anniversary in 2026. The three themes are editorial summaries, not direct quotations.');
	card(slide, '01', 'CREATE VALUE', 'The founding idea was to turn underused resources into valuable materials and industrial activity.', 0.59, 2.24);
	card(slide, '02', 'BUILD CAPABILITY', 'Joint ventures and acquisitions expanded production, technology and access to global markets.', 4.32, 2.24);
	card(slide, '03', 'KEEP EVOLVING', 'Circular-material solutions and the Aramco partnership added new dimensions to SABIC’s journey.', 8.05, 2.24, 3.15);
	text(slide, '1976–2026: fifty years of industrial development and reinvention.', 0.59, 6.14, 10.6, 0.52, { fontSize: 20, color: COLORS.cyan });

	slide = add('OFFICIAL SOURCES', 'Primary sources from sabic.com · accessed 7 October 2026', [1, 2, 3, 4, 5], 'All historical facts and quantitative snapshots in this deck were sourced from SABIC’s own website. Photographs are from the 50 Years page: /LandingPageAssets/50-years/images/1977_img.png and /LandingPageAssets/50-years/images/1983_img.png. Layout branding comes from the user-supplied Template.pdf.');
	for (const [i, [name, url]] of SOURCES.entries()) {
		const y = 2.0 + i * 0.88;
		text(slide, `[${i + 1}] ${name}`, 0.59, y, 10.6, 0.37, { fontSize: 15, bold: true, color: COLORS.cyan, hyperlink: { url } });
		text(slide, url, 0.59, y + 0.44, 10.6, 0.28, { fontSize: 7.7, hyperlink: { url } });
	}

	assert.equal(pptx._slides.length, 12);
	await pptx.writeFile({ fileName: path.join(__dirname, '..', 'SabicHistory.pptx') });
	console.log('Created SabicHistory.pptx: 12 slides, 2 historical photographs, 2 editable charts, and official source citations.');
}

main().catch(error => { console.error(error); process.exitCode = 1; });
