/* How a numbered list counts across its w:num, as Word does it:
   - every w:num of one w:abstractNum continues the same list;
   - a w:num whose w:lvlOverride carries a w:startOverride restarts that level
     once, at the first paragraph the w:num numbers there.
   Checked in both ways the library writes numbers: CSS counters (the default)
   and numberingAsText (the labels as text, e.g. for a PDF). */
describe("numbered lists", function () {
	const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
	const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
	// "name 4" pairs of a computed counter-reset/-increment/-set
	const COUNTER = /([A-Za-z_][\w-]*)\s+(-?\d+)/g;
	// the strings and counter() of a computed content value
	const CONTENT_PART = /"((?:[^"\\]|\\.)*)"|counter\(([\w-]+)[^)]*\)/g;

	// Klarso's Word export of a heading with a fixed number: the headings are
	// numbered by their styles (w:num 1), the fixed one by its own w:num 5000 on
	// the same w:abstractNum, which restarts level 0 at 5.
	// Word shows 1, 1.1, 2, 5, 5.1, 6.
	const headingsWithFixedNumber = {
		numbering: `
			<w:abstractNum w:abstractNumId="0">
				<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="Kapitel"/><w:lvlText w:val="%1"/></w:lvl>
				<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="Abschnitt"/><w:lvlText w:val="%1.%2"/></w:lvl>
			</w:abstractNum>
			<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
			<w:num w:numId="5000"><w:abstractNumId w:val="0"/>
				<w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride>
			</w:num>`,
		styles: `
			<w:style w:type="paragraph" w:styleId="Kapitel"><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr></w:style>
			<w:style w:type="paragraph" w:styleId="Abschnitt"><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr></w:pPr></w:style>`,
		body: [
			paragraph("Kapitel A", { style: "Kapitel" }),
			paragraph("Abschnitt A.1", { style: "Abschnitt" }),
			paragraph("Kapitel B", { style: "Kapitel" }),
			paragraph("Kapitel C", { style: "Kapitel", numId: 5000, ilvl: 0 }),
			paragraph("Abschnitt C.1", { style: "Abschnitt" }),
			paragraph("Kapitel D", { style: "Kapitel" })
		].join("")
	};

	// Word's "Restart Numbering": the list goes on in a new w:num that restarts
	// it at 1, and every item after that uses the new w:num.
	// Word shows 1., 2., 1., 2., 3.
	const restartedList = {
		numbering: `
			<w:abstractNum w:abstractNumId="0">
				<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>
			</w:abstractNum>
			<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
			<w:num w:numId="2"><w:abstractNumId w:val="0"/>
				<w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride>
			</w:num>`,
		body: [
			paragraph("item 1", { numId: 1, ilvl: 0 }),
			paragraph("item 2", { numId: 1, ilvl: 0 }),
			paragraph("restarted item 1", { numId: 2, ilvl: 0 }),
			paragraph("restarted item 2", { numId: 2, ilvl: 0 }),
			paragraph("restarted item 3", { numId: 2, ilvl: 0 })
		].join("")
	};

	let container = null;

	afterEach(() => container?.remove());

	it("go on from a heading with a fixed number (CSS counters)", async () => {
		container = await render(headingsWithFixedNumber);

		expect(cssCounterLabels(container)).toEqual(["1", "1.1", "2", "5", "5.1", "6"]);
	});

	it("go on from a heading with a fixed number (numberingAsText)", async () => {
		container = await render(headingsWithFixedNumber, { numberingAsText: true });

		expect(textLabels(container)).toEqual(["1", "1.1", "2", "5", "5.1", "6"]);
	});

	it("restart once, at the first item of the restarting w:num (CSS counters)", async () => {
		container = await render(restartedList);

		expect(cssCounterLabels(container)).toEqual(["1.", "2.", "1.", "2.", "3."]);
	});

	it("restart once, at the first item of the restarting w:num (numberingAsText)", async () => {
		container = await render(restartedList, { numberingAsText: true });

		expect(textLabels(container)).toEqual(["1.", "2.", "1.", "2.", "3."]);
	});

	// a paragraph with a style and/or numbering of its own
	function paragraph(text, { style, numId, ilvl } = {}) {
		const props = (style ? `<w:pStyle w:val="${style}"/>` : "")
			+ (numId != null ? `<w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr>` : "");
		return `<w:p><w:pPr>${props}</w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
	}

	// renders a minimal .docx with the given numbering, styles and body
	async function render({ numbering, styles = "", body }, options) {
		const zip = new JSZip();
		zip.file("[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
			<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
			<Default Extension="xml" ContentType="application/xml"/>
			<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
		</Types>`);
		zip.file("_rels/.rels", relationships(`<Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/>`));
		zip.file("word/_rels/document.xml.rels", relationships(
			`<Relationship Id="rId1" Type="${REL}/styles" Target="styles.xml"/>`
			+ `<Relationship Id="rId2" Type="${REL}/numbering" Target="numbering.xml"/>`));
		zip.file("word/styles.xml", `<w:styles ${W}>${styles}</w:styles>`);
		zip.file("word/numbering.xml", `<w:numbering ${W}>${numbering}</w:numbering>`);
		zip.file("word/document.xml", `<w:document ${W}><w:body>${body}
			<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1134" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>
		</w:body></w:document>`);

		const result = document.createElement("div");
		document.body.appendChild(result);
		await docx.renderAsync(await zip.generateAsync({ type: "blob" }), result, null, options);
		return result;
	}

	function relationships(content) {
		return `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${content}</Relationships>`;
	}

	// numberingAsText: the label span each numbered paragraph starts with
	function textLabels(container) {
		return Array.from(container.querySelectorAll("[data-docx-num-label]"), span => span.textContent.trim());
	}

	/* CSS counters: the ::before content of each numbered paragraph, with the
	   counters counted as the browser counts them, in tree order: on every
	   element counter-reset, then counter-increment, then counter-set, and the
	   same on its ::before. Decimal counters only. */
	function cssCounterLabels(container) {
		const counters = {};
		const count = (style) => {
			for (const [, name, value] of style.counterReset.matchAll(COUNTER))
				counters[name] = Number(value);
			for (const [, name, value] of style.counterIncrement.matchAll(COUNTER))
				counters[name] = (counters[name] ?? 0) + Number(value);
			for (const [, name, value] of style.counterSet.matchAll(COUNTER))
				counters[name] = Number(value);
		};

		const labels = [];
		for (const element of [container, ...container.querySelectorAll("*")]) {
			count(getComputedStyle(element));
			const before = getComputedStyle(element, "::before");
			count(before);
			if (before.content.includes("counter("))
				labels.push(Array.from(before.content.matchAll(CONTENT_PART),
					([, text, counter]) => counter ? counters[counter] : unescapeCss(text)).join("").trim());
		}
		return labels;
	}

	function unescapeCss(text) {
		return text.replace(/\\([0-9a-fA-F]{1,6})\s?/g, (s, hex) => String.fromCodePoint(parseInt(hex, 16)))
			.replace(/\\(.)/g, "$1");
	}
});
