/* A page held at the paper's height - measured pagination pins every sheet
   (pinSheetHeights) - whose body is too tall for it: the body runs on into the
   bottom margin, the header and the footer stay where they are on a page that
   fits. The page is a flex column, and a header or footer that may shrink to
   its reserved band lets the body slide up over the header. */
describe("an overfull page", function () {
	const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
	const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
	const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

	let container = null;

	afterEach(() => container?.remove());

	it("keeps the body below the whole header", async () => {
		container = await render();
		const page = container.querySelector("section.docx");
		const body = page.querySelector(":scope > article");
		const bodyTopOnFittingPage = body.getBoundingClientRect().top;

		overfill(page);

		expect(body.getBoundingClientRect().top).toBeCloseTo(bodyTopOnFittingPage, 0);
	});

	it("keeps the footer where a fitting page has it", async () => {
		container = await render();
		const page = container.querySelector("section.docx");
		const footer = page.querySelector(":scope > footer");
		const footerBottomOnFittingPage = footer.getBoundingClientRect().bottom;

		overfill(page);

		expect(footer.getBoundingClientRect().bottom).toBeCloseTo(footerBottomOnFittingPage, 0);
	});

	// the page as measured pagination leaves it, with more body than it holds
	function overfill(page) {
		page.style.height = getComputedStyle(page).minHeight;
		const tooMuch = document.createElement("div");
		tooMuch.style.height = page.getBoundingClientRect().height + "px";
		page.querySelector(":scope > article").appendChild(tooMuch);
	}

	function paragraph(text) {
		return `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
	}

	// a page with a 1 cm header band (top margin 2 cm, header at 1 cm) and a
	// five line header - taller than its band, so it pushes the body down
	async function render() {
		const zip = new JSZip();
		zip.file("[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
			<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
			<Default Extension="xml" ContentType="application/xml"/>
			<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
			<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
			<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
		</Types>`);
		zip.file("_rels/.rels", relationships(`<Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/>`));
		zip.file("word/_rels/document.xml.rels", relationships(
			`<Relationship Id="rIdH" Type="${REL}/header" Target="header1.xml"/>`
			+ `<Relationship Id="rIdF" Type="${REL}/footer" Target="footer1.xml"/>`));
		zip.file("word/header1.xml", `<w:hdr ${W}>${["Kopf 1", "Kopf 2", "Kopf 3", "Kopf 4", "Kopf 5"].map(paragraph).join("")}</w:hdr>`);
		zip.file("word/footer1.xml", `<w:ftr ${W}>${paragraph("Fuß")}</w:ftr>`);
		zip.file("word/document.xml", `<w:document ${W} ${R}><w:body>
			${paragraph("Erster Absatz")}${paragraph("Zweiter Absatz")}
			<w:sectPr>
				<w:headerReference w:type="default" r:id="rIdH"/><w:footerReference w:type="default" r:id="rIdF"/>
				<w:pgSz w:w="11906" w:h="16838"/>
				<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="567" w:footer="567" w:gutter="0"/>
			</w:sectPr>
		</w:body></w:document>`);

		const result = document.createElement("div");
		document.body.appendChild(result);
		await docx.renderAsync(await zip.generateAsync({ type: "blob" }), result);
		return result;
	}

	function relationships(content) {
		return `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${content}</Relationships>`;
	}
});
