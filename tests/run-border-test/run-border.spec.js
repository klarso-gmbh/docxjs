/* A run's border (w:rPr/w:bdr) goes around everything the run holds. The
   export puts an image with a frame into a run of its own and borders that
   run, so the border has to go around the whole picture - not only around
   the line of text the picture sits on. A bordered run of text stays text:
   it still breaks across lines. */
describe("a bordered run", function () {
	const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
	const DRAWING = 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
		+ ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
		+ ' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'
		+ ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
	const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
	// a 1 x 1 px PNG
	const PNG = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";
	// the border the export gives an image's run: 12 eighths of a point = 1.5pt = 2px, no gap
	const BORDER = `<w:rPr><w:bdr w:val="single" w:sz="12" w:space="0" w:color="000000"/></w:rPr>`;

	let container = null;

	afterEach(() => container?.remove());

	it("draws its border around the whole picture it holds", async () => {
		container = await render(`<w:r><w:t xml:space="preserve">davor </w:t></w:r>${pictureRun(BORDER)}<w:r><w:t xml:space="preserve"> danach</w:t></w:r>`);
		const picture = container.querySelector("img").getBoundingClientRect();
		const run = container.querySelector("img").closest("span").getBoundingClientRect();

		expect([run.left, run.top, run.right, run.bottom])
			.toEqual([picture.left - 2, picture.top - 2, picture.right + 2, picture.bottom + 2]);
	});

	it("still breaks across lines when it holds text", async () => {
		container = await render(`<w:r>${BORDER}<w:t>${"Rahmentext ".repeat(40)}</w:t></w:r>`);

		expect(container.querySelector("article span").getClientRects().length).toBeGreaterThan(1);
	});

	// a run holding an inline picture of 200 x 100 px (1905000 x 952500 EMU), with `props` as its w:rPr
	function pictureRun(props) {
		return `<w:r>${props}<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">
			<wp:extent cx="1905000" cy="952500"/><wp:docPr id="1" name="Bild"/>
			<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>
				<pic:nvPicPr><pic:cNvPr id="1" name="Bild"/><pic:cNvPicPr/></pic:nvPicPr>
				<pic:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
				<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="952500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
			</pic:pic></a:graphicData></a:graphic>
		</wp:inline></w:drawing></w:r>`;
	}

	// renders a minimal .docx: one paragraph holding `runs`, and the image they show
	async function render(runs) {
		const zip = new JSZip();
		zip.file("[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
			<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
			<Default Extension="xml" ContentType="application/xml"/>
			<Default Extension="png" ContentType="image/png"/>
			<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
		</Types>`);
		zip.file("_rels/.rels", relationships(`<Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/>`));
		zip.file("word/_rels/document.xml.rels", relationships(`<Relationship Id="rIdImage" Type="${REL}/image" Target="media/image1.png"/>`));
		zip.file("word/media/image1.png", PNG, { base64: true });
		zip.file("word/document.xml", `<w:document ${W} ${DRAWING}><w:body><w:p>${runs}</w:p>
			<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="567" w:footer="567" w:gutter="0"/></w:sectPr>
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
