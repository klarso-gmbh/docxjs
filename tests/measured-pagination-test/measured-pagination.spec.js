/* Measured pagination keeps the pages a document's saved page breaks
   (w:lastRenderedPageBreak - Word's own layout) delimit while such a page
   overflows by less than a tolerance: the browser sets text a little longer
   than Word does. A document without saved breaks - written by a program, not
   by Word - has no such pages. Every page comes from the measurement, and a
   page one line too long is broken where Word breaks it. */
describe("measured pagination", function () {
	const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
	const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
	// as Suite's preview and PDF pages render: measured, the saved breaks kept
	const OPTIONS = { measuredPagination: true, ignoreLastRenderedPageBreak: false };

	let container = null;

	afterEach(() => container?.remove());

	// the body band of the page below: 16838 - 2 x 1134 twips = 971.3 px;
	// 49 lines of exactly 20 px take 980 px - one line too many
	it("breaks a document without saved page breaks at the line that does not fit", async () => {
		container = await render(lines(49));

		expect(container.querySelectorAll("section.docx").length).toBe(2);
	});

	it("keeps a document without saved page breaks that fits on one page", async () => {
		container = await render(lines(48));

		expect(container.querySelectorAll("section.docx").length).toBe(1);
	});

	it("keeps an explicit page break of a document without saved page breaks", async () => {
		container = await render(lines(10) + pageBreak() + lines(10));

		expect(container.querySelectorAll("section.docx").length).toBe(2);
	});

	// what does not fit before an explicit page break gets a page of its own,
	// and the break then starts the next one - as in Word
	it("moves the lines that do not fit before an explicit page break to a page of their own", async () => {
		container = await render(lines(49) + pageBreak() + lines(1));

		expect(container.querySelectorAll("section.docx").length).toBe(3);
	});

	const EXACT_LINE = `<w:spacing w:before="0" w:after="0" w:line="300" w:lineRule="exact"/>`;

	// paragraphs of one line, exactly 15 pt (20 px) high, no space around them
	function lines(count) {
		let xml = "";
		for (let i = 1; i <= count; i++)
			xml += `<w:p><w:pPr>${EXACT_LINE}</w:pPr><w:r><w:t>Zeile ${i}</w:t></w:r></w:p>`;
		return xml;
	}

	// a paragraph holding nothing but a page break (Ctrl+Enter in Word)
	function pageBreak() {
		return `<w:p><w:pPr>${EXACT_LINE}</w:pPr><w:r><w:br w:type="page"/></w:r></w:p>`;
	}

	async function render(body) {
		const zip = new JSZip();
		zip.file("[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
			<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
			<Default Extension="xml" ContentType="application/xml"/>
			<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
		</Types>`);
		zip.file("_rels/.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`);
		zip.file("word/document.xml", `<w:document ${W}><w:body>${body}
			<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="567" w:footer="567" w:gutter="0"/></w:sectPr>
		</w:body></w:document>`);

		const result = document.createElement("div");
		document.body.appendChild(result);
		await docx.renderAsync(await zip.generateAsync({ type: "blob" }), result, null, OPTIONS);
		return result;
	}
});
