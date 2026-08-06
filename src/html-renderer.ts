import { WordDocument } from './word-document';
import {
	DomType, WmlTable, IDomNumbering,
	WmlHyperlink, IDomImage, OpenXmlElement, WmlTableColumn, WmlTableCell, WmlText, WmlSymbol, WmlBreak, WmlNoteReference,
	WmlSmartTag,
	WmlAltChunk,
	WmlTableRow,
	WmlChange
} from './document/dom';
import { Options } from './docx-preview';
import { DocumentElement } from './document/document';
import { WmlParagraph } from './document/paragraph';
import { asArray, encloseFontFamily, escapeClassName, isString, keyBy, mergeDeep } from './utils';
import { computePixelToPoint, updateTabStop } from './javascript';
import { FontTablePart } from './font-table/font-table';
import { FooterHeaderReference, SectionProperties } from './document/section';
import { WmlRun } from './document/run';
import { WmlFieldSimple } from './document/fields';
import { WmlBookmarkStart } from './document/bookmarks';
import { IDomStyle } from './document/style';
import { WmlBaseNote, WmlFootnote } from './notes/elements';
import { ThemePart } from './theme/theme-part';
import { BaseHeaderFooterPart } from './header-footer/parts';
import { Part } from './common/part';
import { VmlElement } from './vml/vml';
import { WmlCommentRangeStart, WmlCommentReference } from './comments/elements';
import { cx, h, ns } from './html';
import {
	assignModelIndices, applyMeasuredGeometry, findMeasuredBreaks, applyBreaksToModel,
	resolveBreaksToDocument, hyphenateModel, getAutoLineMul
} from './measured-pagination';

interface CellPos {
	col: number;
	row: number;
}

interface Section {
	sectProps: SectionProperties;
	elements: OpenXmlElement[];
	pageBreak: boolean;
}

declare const Highlight: any;

type CellVerticalMergeType = Record<number, HTMLTableCellElement>;

export class HtmlRenderer {

	className: string = "docx";
	rootSelector: string;
	document: WordDocument;
	options: Options;
	styleMap: Record<string, IDomStyle> = {};
	currentPart: Part = null;

	tableVerticalMerges: CellVerticalMergeType[] = [];
	currentVerticalMerge: CellVerticalMergeType = null;
	tableCellPositions: CellPos[] = [];
	currentCellPosition: CellPos = null;

	footnoteMap: Record<string, WmlFootnote> = {};
	endnoteMap: Record<string, WmlFootnote> = {};
	currentFootnoteIds: string[];
	currentEndnoteIds: string[] = [];
	usedHederFooterParts: any[] = [];

	// measured pagination: mi -> body element map, current page ordinal for
	// mirror margins, and rendered header/footer part roots by path (for
	// applying measured geometry to the part model that the second pass re-uses)
	measuredMiMap: Map<number, OpenXmlElement> = null;
	currentPageIndex = 0;
	hfPartsByPath: Record<string, OpenXmlElement> = {};

	defaultTabSize: string;
	currentTabs: any[] = [];

	commentHighlight: any;
	commentMap: Record<string, Range> = {};

	tasks: Promise<any>[] = [];
	postRenderTasks: any[] = [];
	h = h;

	async render(document: WordDocument, options: Options): Promise<Node[]> {
		this.document = document;
		this.options = options;
		this.className = options.className;
		this.rootSelector = options.inWrapper ? `.${this.className}-wrapper` : ':root';
		this.h = options.h ?? h;
		this.styleMap = null;
		this.tasks = [];

		if (this.options.renderComments && globalThis.Highlight) {
			this.commentHighlight = new Highlight();
		}

		const result: Node[] = [...this.renderDefaultStyle()];

		if (document.themePart) {
			result.push(...this.renderTheme(document.themePart));
		}

		if (document.stylesPart != null) {
			this.styleMap = this.processStyles(document.stylesPart.styles);
			result.push(...this.renderStyles(document.stylesPart.styles));
		}

		if (document.numberingPart) {
			this.prodessNumberings(document.numberingPart.domNumberings);

			result.push(...await this.renderNumbering(document.numberingPart.domNumberings));
			//result.push(...await this.renderNumbering2(document.numberingPart.domNumberings));
		}

		if (document.footnotesPart) {
			this.footnoteMap = keyBy(document.footnotesPart.notes, x => x.id);
		}

		if (document.endnotesPart) {
			this.endnoteMap = keyBy(document.endnotesPart.notes, x => x.id);
		}

		if (document.settingsPart) {
			this.defaultTabSize = document.settingsPart.settings?.defaultTabStop;
		}

		if (!options.ignoreFonts && document.fontTablePart)
			result.push(...await this.renderFontTable(document.fontTablePart));

		var sectionElements = options.measuredPagination
			? await this.renderMeasured(document, result)
			: this.renderSections(document.documentPart.body);

		if (this.options.inWrapper) {
			result.push(this.renderWrapper(sectionElements));
		} else {
			result.push(...sectionElements);
		}

		if (this.commentHighlight && options.renderComments) {
			(CSS as any).highlights.set(`${this.className}-comments`, this.commentHighlight);
		}

		this.postRenderTasks.forEach(t => t());

		await Promise.allSettled(this.tasks);

		this.refreshTabStops();

		return result;
	}

	renderTheme(themePart: ThemePart) {
		const variables = {};
		const fontScheme = themePart.theme?.fontScheme;

		if (fontScheme) {
			if (fontScheme.majorFont) {
				variables['--docx-majorHAnsi-font'] = fontScheme.majorFont.latinTypeface;
			}

			if (fontScheme.minorFont) {
				variables['--docx-minorHAnsi-font'] = fontScheme.minorFont.latinTypeface;
			}
		}

		const colorScheme = themePart.theme?.colorScheme;

		if (colorScheme) {
			for (let [k, v] of Object.entries(colorScheme.colors)) {
				variables[`--docx-${k}-color`] = `#${v}`;
			}
		}

		const cssText = this.styleToString(`.${this.className}`, variables);
		return [
			this.h({ tagName: "#comment", children: ["docxjs document theme values"] }),
			this.h({ tagName: "style", children: [cssText] })
		];
	}

	async renderFontTable(fontsPart: FontTablePart) {
		const result = [];

		for (let f of fontsPart.fonts) {
			for (let ref of f.embedFontRefs) {
				try{
					const fontData = await this.document.loadFont(ref.id, ref.key);
					const cssValues = {
						'font-family': encloseFontFamily(f.name),
						'src': `url(${fontData})`
					};

					if (ref.type == "bold" || ref.type == "boldItalic") {
						cssValues['font-weight'] = 'bold';
					}

					if (ref.type == "italic" || ref.type == "boldItalic") {
						cssValues['font-style'] = 'italic';
					}

					result.push(this.h({ tagName: "#comment", children: [`docxjs ${f.name} font`] }));
					result.push(this.h({ tagName: "style", children: [this.styleToString(`@font-face`, cssValues)] }));
				} catch(e) {
					if (this.options.debug) console.warn(`Can't load font with id ${ref.id} and key ${ref.key}`);
				}
			}
		}

		return result;
	}

	processStyleName(className: string): string {
		return className ? `${this.className}_${escapeClassName(className)}` : this.className;
	}

	processStyles(styles: IDomStyle[]) {
		const stylesMap = keyBy(styles.filter(x => x.id != null), x => x.id);

		for (const style of styles.filter(x => x.basedOn)) {
			var baseStyle = stylesMap[style.basedOn];

			if (baseStyle) {
				style.paragraphProps = mergeDeep(style.paragraphProps, baseStyle.paragraphProps);
				style.runProps = mergeDeep(style.runProps, baseStyle.runProps);

				for (const baseValues of baseStyle.styles) {
					const styleValues = style.styles.find(x => x.target == baseValues.target);

					if (styleValues) {
						this.copyStyleProperties(baseValues.values, styleValues.values);
					} else {
						style.styles.push({ ...baseValues, values: { ...baseValues.values } });
					}
				}
			}
			else if (this.options.debug)
				console.warn(`Can't find base style ${style.basedOn}`);
		}

		for (let style of styles) {
			style.cssName = this.processStyleName(style.id);
		}

		return stylesMap;
	}

	prodessNumberings(numberings: IDomNumbering[]) {
		for (let num of numberings.filter(n => n.pStyleName)) {
			const style = this.findStyle(num.pStyleName);

			if (style?.paragraphProps?.numbering) {
				style.paragraphProps.numbering.level = num.level;
			}
		}
	}

	processElement(element: OpenXmlElement) {
		if (element.children) {
			for (var e of element.children) {
				e.parent = element;

				if (e.type == DomType.Table) {
					this.processTable(e);
				}
				else {
					this.processElement(e);
				}
			}
		}
	}

	processTable(table: WmlTable) {
		for (var r of table.children) {
			for (var c of r.children) {
				c.cssStyle = this.copyStyleProperties(table.cellStyle, c.cssStyle, [
					"border-left", "border-right", "border-top", "border-bottom",
					"padding-left", "padding-right", "padding-top", "padding-bottom"
				]);

				this.processElement(c);
			}
		}
	}

	copyStyleProperties(input: Record<string, string>, output: Record<string, string>, attrs: string[] = null): Record<string, string> {
		if (!input)
			return output;

		if (output == null) output = {};
		if (attrs == null) attrs = Object.getOwnPropertyNames(input);

		for (var key of attrs) {
			if (input.hasOwnProperty(key) && !output.hasOwnProperty(key))
				output[key] = input[key];
		}

		return output;
	}

	createPageElement(className: string, props: SectionProperties, docStyle: Record<string, any>) {
		const style: Record<string, string> = { ...docStyle };

		if (props) {
			if (props.pageMargins) {
				style.paddingLeft = props.pageMargins.left;
				style.paddingRight = props.pageMargins.right;
				style.paddingTop = props.pageMargins.top;
				style.paddingBottom = props.pageMargins.bottom;
			}

			if (props.pageSize) {
				if (!this.options.ignoreWidth)
					style.width = props.pageSize.width;
				if (!this.options.ignoreHeight)
					style.minHeight = props.pageSize.height;
			}

			// mirrored margins: even pages swap left/right page padding (content
			// width is parity-constant, so this is a pure x-shift)
			if (this.options.measuredPagination && props.pageMargins
				&& this.document.settingsPart?.settings?.mirrorMargins
				&& (this.currentPageIndex + 1) % 2 === 0) {
				const l = style.paddingLeft;
				style.paddingLeft = style.paddingRight;
				style.paddingRight = l;
			}
		}

		return this.h({ tagName: "section", className, style }) as HTMLElement;
	}

	createSectionContent(props: SectionProperties) {
		const style: Record<string, string> = {};

		if (props.columns && props.columns.numberOfColumns) {
			style.columnCount = `${props.columns.numberOfColumns}`;
			style.columnGap = props.columns.space;

			if (props.columns.separator) {
				style.columnRule = "1px solid black";
			}
		}

		return this.h({ tagName: "article", style });
	}	

	renderSections(document: DocumentElement): HTMLElement[] {
		const result = [];

		this.processElement(document);
		const sections = this.splitBySection(document.children, document.props);
		const pages = this.groupByPageBreaks(sections);
		let prevProps = null;
		const sectIndex = new Map<SectionProperties, number>();

		for (let i = 0, l = pages.length; i < l; i++) {
			this.currentFootnoteIds = [];

			const section = pages[i][0];
			let props = section.sectProps;
			this.currentPageIndex = result.length; // global page ordinal (mirror margins)
			const pageElement = this.createPageElement(this.className, props, document.cssStyle);

			// stamp the Word section index (pages of one section share sectProps)
			// so a consumer's page-number stamper can count SECTIONPAGES
			if (this.options.measuredPagination) {
				if (!sectIndex.has(props)) sectIndex.set(props, sectIndex.size);
				pageElement.setAttribute("data-docx-sect", String(sectIndex.get(props)));
			}

			this.options.renderHeaders && this.renderHeaderFooter(props.headerRefs, props,
				result.length, prevProps != props, pageElement);

			for (const sect of pages[i]) {
				var contentElement = this.createSectionContent(sect.sectProps);
				this.renderElements(sect.elements, contentElement);
				pageElement.appendChild(contentElement);
				props = sect.sectProps;
			}

			if (this.options.renderFootnotes) {
				const notes = this.renderNotes(this.currentFootnoteIds, this.footnoteMap);
				notes && pageElement.appendChild(notes);
			}

			if (this.options.renderEndnotes && i == l - 1) {
				const notes = this.renderNotes(this.currentEndnoteIds, this.endnoteMap);
				notes && pageElement.appendChild(notes);
			}

			this.options.renderFooters && this.renderHeaderFooter(props.footerRefs, props,
				result.length, prevProps != props, pageElement);

			result.push(pageElement);
			prevProps = props;
		}

		return result;
	}

	/* Two-pass measured pagination: render once, measure where content overflows
	   each page, insert lastRenderedPageBreak markers into the model, then render
	   again so pages split natively - with per-page header/footer selection and
	   mirrored margins handled by the normal render path. Word "auto" line
	   metrics, justified space shrinking and hyphenation are applied to the model
	   before measuring so both passes agree. */
	async renderMeasured(wordDoc: WordDocument, styleResult: Node[]): Promise<HTMLElement[]> {
		const body = wordDoc.documentPart.body;
		const settings = wordDoc.settingsPart?.settings;

		// measured pagination expresses breaks as lastRenderedPageBreak markers,
		// so both passes must honor them (and pass 1 must group the doc into as few
		// sheets as its own breaks dictate, then measure the overflow)
		this.options.ignoreLastRenderedPageBreak = false;

		// hyphenate the model text (both passes see the same soft hyphens)
		if (this.options.hyphenate && settings?.autoHyphenation) {
			try { await hyphenateModel(body, this.options.hyphenate); }
			catch (e) { if (this.options.debug) console.warn("docx-preview: hyphenation failed", e); }
		}

		// index body paragraphs/tables so measured DOM maps back to the model
		this.measuredMiMap = assignModelIndices(body);

		// pass 1: render every header/footer variant so continuation sheets are
		// measured with their real (default) header, not the tall first-page one
		this.options.renderHeaderFooterVariants = true;
		this.hfPartsByPath = {};
		const pass1 = this.renderSections(body);
		// pass 2 selects one variant per page natively (by final ordinal), so it
		// must not keep every variant around
		this.options.renderHeaderFooterVariants = false;

		// offscreen measuring host with the document's own styles applied
		const host = this.h({ tagName: "div", className: `${this.className}-wrapper`, children: pass1 }) as HTMLElement;
		host.style.cssText = "position:absolute;left:-99999px;top:0;visibility:hidden";
		const styleClones = styleResult.filter(n => n.nodeName === "STYLE").map(n => n.cloneNode(true) as Node);
		styleClones.forEach(s => document.head.appendChild(s));
		document.body.appendChild(host);

		try {
			// image boxes are sized from the model, but await loads to match on-screen
			await Promise.allSettled(this.tasks);

			const lineMul = getAutoLineMul(wordDoc);
			const shrink = this.options.justificationShrink != null
				? -Math.abs(this.options.justificationShrink) : undefined;
			applyMeasuredGeometry(host, this.measuredMiMap, this.hfPartsByPath, settings, lineMul, shrink);

			const breaks = findMeasuredBreaks(host);
			// expose the breaks addressed against the original document BEFORE
			// applyBreaksToModel mutates the paragraphs (offsets still line up)
			if (this.options.onMeasuredBreaks) {
				try {
					this.options.onMeasuredBreaks(resolveBreaksToDocument(breaks, this.measuredMiMap, body));
				} catch (e) {
					if (this.options.debug) console.warn("docx-preview: onMeasuredBreaks callback failed", e);
				}
			}
			applyBreaksToModel(breaks, this.measuredMiMap);
		} finally {
			host.remove();
			styleClones.forEach(s => (s as ChildNode).remove());
		}

		// reset accumulated render state for the clean second pass
		this.currentFootnoteIds = [];
		this.currentEndnoteIds = [];
		this.usedHederFooterParts = [];
		this.tableVerticalMerges = [];
		this.currentVerticalMerge = null;
		this.tableCellPositions = [];
		this.currentCellPosition = null;
		this.tasks = [];
		this.postRenderTasks = [];

		// pass 2: render the now-marked model (native per-page header/footer + mirror)
		this.currentPageIndex = 0;
		const pages = this.renderSections(body);
		this.fillPageFields(pages);
		return pages;
	}

	renderHeaderFooter(refs: FooterHeaderReference[], props: SectionProperties, page: number, firstOfSection: boolean, into: HTMLElement) {
		if (!refs) return;

		// even/odd headers only apply when the document opts in (Word behavior)
		const evenAndOdd = this.document.settingsPart?.settings?.evenAndOddHeaders ?? false;

		var ref = (props.titlePage && firstOfSection ? refs.find(x => x.type == "first") : null)
			?? (evenAndOdd && page % 2 == 1 ? refs.find(x => x.type == "even") : null)
			?? refs.find(x => x.type == "default");

		// with renderHeaderFooterVariants every variant a later page of this
		// section could need is rendered, tagged with data-hf-type and hidden
		// unless selected - consumers that re-paginate by height (cloning
		// sections into more pages) can then re-select per page
		var variants = this.options.renderHeaderFooterVariants
			? refs.filter((x, i) => (x.type == "default" || x.type == null
					|| (x.type == "first" && props.titlePage)
					|| (x.type == "even" && evenAndOdd))
				&& refs.findIndex(y => y.type == x.type) == i)
			: (ref ? [ref] : []);

		for (const variant of variants) {
			const part = this.document.findPartByRelId(variant.id, this.document.documentPart) as BaseHeaderFooterPart;

			if (!part) continue;

			this.currentPart = part;
			if (!this.usedHederFooterParts.includes(part.path)) {
				this.processElement(part.rootElement);
				this.usedHederFooterParts.push(part.path);
			}
			const [el] = this.renderElements([part.rootElement], into) as HTMLElement[];

			if (this.options.renderHeaderFooterVariants) {
				el.setAttribute("data-hf-type", variant.type ?? "default");
				if (variant !== ref) el.style.display = "none";
			}

			// measured pagination applies Word line metrics to the header/footer
			// part model too; tag the instance so geometry maps its <p> back
			if (this.options.measuredPagination) {
				el.setAttribute("data-docx-part", part.path);
				this.hfPartsByPath[part.path] = part.rootElement;
			}

			if (props?.pageMargins) {
				if (part.rootElement.type === DomType.Header) {
					el.style.marginTop = `calc(${props.pageMargins.header} - ${props.pageMargins.top})`;
					el.style.minHeight = `calc(${props.pageMargins.top} - ${props.pageMargins.header})`;
				}
				else if (part.rootElement.type === DomType.Footer) {
					el.style.marginBottom = `calc(${props.pageMargins.footer} - ${props.pageMargins.bottom})`;
					el.style.minHeight = `calc(${props.pageMargins.bottom} - ${props.pageMargins.footer})`;
					// footers grow upward from the footer line: anchor the
					// content at the bottom of the reserved band, like Word
					el.style.display = "flex";
					el.style.flexDirection = "column";
					el.style.justifyContent = "flex-end";
				}
			}
		}

		this.currentPart = null;
	}

	isPageBreakElement(elem: OpenXmlElement): boolean {
		if (elem.type != DomType.Break)
			return false;

		if ((elem as WmlBreak).break == "lastRenderedPageBreak")
			return !this.options.ignoreLastRenderedPageBreak;

		return (elem as WmlBreak).break == "page";
	}

	// a run whose direct children include an honored page break
	isPageBreakRun(elem: OpenXmlElement): boolean {
		return elem?.type == DomType.Run && (elem.children?.some(c => this.isPageBreakElement(c)) ?? false);
	}

	isPageBreakSection(prev: SectionProperties, next: SectionProperties): boolean {
		if (!prev) return false;
		if (!next) return false;

		return prev.pageSize?.orientation != next.pageSize?.orientation
			|| prev.pageSize?.width != next.pageSize?.width
			|| prev.pageSize?.height != next.pageSize?.height;
	}

	// Scans a table row for a page break that splits the table there. Only the
	// lastRenderedPageBreak markers measured pagination inserts count: Word does
	// NOT break a page at a MANUAL break inside a table cell - it lays the table
	// out as if the break were not there (verified against Word's PDF for a
	// document whose tables carry a trailing <w:br w:type="page"/> in their last
	// cell) - so honoring those would tear a table apart where Word keeps it.
	rowHasPageBreak(row: OpenXmlElement): boolean {
		const scan = (el: OpenXmlElement): boolean =>
			(el.type == DomType.Break && (el as WmlBreak).break == "lastRenderedPageBreak"
				&& !this.options.ignoreLastRenderedPageBreak)
			|| (el.children?.some(scan) ?? false);
		return row.children?.some(scan) ?? false;
	}

	// splits a table into row segments at cell-internal page breaks; every
	// segment is a clone keeping the columns (widths). Returns [table] when there
	// is nothing to split.
	splitTableByRowBreaks(table: WmlTable): WmlTable[] {
		const rows = table.children ?? [];
		const breakRows: number[] = [];
		for (let ri = 0; ri < rows.length; ri++)
			if (this.rowHasPageBreak(rows[ri])) breakRows.push(ri);
		if (!breakRows.length) return [table];

		const bounds = [0, ...breakRows, rows.length];
		const segs: WmlTable[] = [];
		for (let bi = 0; bi < bounds.length - 1; bi++) {
			const seg = rows.slice(bounds[bi], bounds[bi + 1]);
			if (!seg.length) continue;
			// Every segment is a CLONE - assigning the first one back to
			// table.children would drop the rows behind the break from the model,
			// and measured pagination renders the same model twice: the second
			// pass would neither split the table nor even show those rows.
			// $rowbase keeps the measured pass addressing rows of the ORIGINAL
			// table, whose $mi all segments share.
			segs.push({ ...table, children: seg, $rowbase: bounds[bi] } as WmlTable);
		}
		return segs;
	}

	splitBySection(elements: OpenXmlElement[], defaultProps: SectionProperties): Section[] {
		var current: Section = { sectProps: null, elements: [], pageBreak: false };
		var result = [current];

		// a paragraph may carry several page breaks (measured pagination inserts
		// them; a Word/marked document can too - a figure whose images span
		// pages, or a break-before plus an internal break). Iterate a worklist so
		// a split continuation can be re-scanned for its own further breaks.
		const measured = this.options.measuredPagination;
		const list = elements.slice();

		for (let ei = 0; ei < list.length; ei++) {
			let elem = list[ei];

			// page break before a block (measured: whole-block moves and table
			// continuations moved as a unit)
			if (measured && (elem as any).$pageBreakBefore) {
				current.pageBreak = true;
				current = { sectProps: null, elements: [], pageBreak: false };
				result.push(current);
			}

			if (elem.type == DomType.Paragraph && !(elem as any).$splitCont) {
				const s = this.findStyle((elem as WmlParagraph).styleName);

				if (s?.paragraphProps?.pageBreakBefore) {
					current.sectProps = sectProps;
					current.pageBreak = true;
					current = { sectProps: null, elements: [], pageBreak: false };
					result.push(current);
				}
			}

			// measured: split a table at cell-internal page breaks (row boundaries)
			if (measured && elem.type == DomType.Table && this.options.breakPages) {
				const segs = this.splitTableByRowBreaks(elem as WmlTable);
				if (segs.length > 1) {
					for (let si = 0; si < segs.length; si++) {
						current.elements.push(segs[si]);
						if (si < segs.length - 1) {
							current.pageBreak = true;
							current = { sectProps: null, elements: [], pageBreak: false };
							result.push(current);
						}
					}
					continue;
				}
			}

			current.elements.push(elem);

			if (elem.type == DomType.Paragraph) {
				const p = elem as WmlParagraph;
				const owner = current;   // the section elem was just pushed into

				var sectProps = p.sectionProps;
				var pBreakIndex = -1;
				var rBreakIndex = -1;

				if (this.options.breakPages && p.children) {
					pBreakIndex = p.children.findIndex(r => {
						rBreakIndex = r.children?.findIndex(this.isPageBreakElement.bind(this)) ?? -1;
						return rBreakIndex != -1;
					});
				}

				if (sectProps || pBreakIndex != -1) {
					current.sectProps = sectProps;
					current.pageBreak = pBreakIndex != -1;
					current = { sectProps: null, elements: [], pageBreak: false };
					result.push(current);
				}

				if (pBreakIndex != -1) {
					let breakRun = p.children[pBreakIndex];
					let splitRun = rBreakIndex < breakRun.children.length - 1;

					if (pBreakIndex < p.children.length - 1 || splitRun) {
						const children = elem.children;
						// Split into two CLONES and swap the first one in for the
						// element already collected. Mutating the model instead would
						// consume the break: measured pagination renders the same
						// model twice, and by the second pass the paragraph no longer
						// carries the marker - the page silently stops breaking there.
						const head = { ...elem, children: children.slice(0, pBreakIndex) };
						var newParagraph = { ...elem, children: children.slice(pBreakIndex) };

						if (splitRun) {
							const runChildren = breakRun.children;
							head.children = head.children.concat([
								{ ...breakRun, children: runChildren.slice(0, rBreakIndex) }]);
							newParagraph.children = ([
								{ ...breakRun, children: runChildren.slice(rBreakIndex) }
							] as OpenXmlElement[]).concat(children.slice(pBreakIndex + 1));
						}

						owner.elements[owner.elements.length - 1] = head;

						// re-inject the continuation so it is re-scanned for its own
						// further breaks (a continuation must not repeat the number,
						// and a break-before belonged to the first part only). This
						// is what lets a paragraph split at MORE THAN ONE marker -
						// needed both under measured pagination and when re-rendering
						// a marked document with measuredPagination off.
						(newParagraph as any).$splitCont = true;
						(newParagraph as any).$pageBreakBefore = false;
						// drop the leading page-break marker so the same break is not
						// re-detected (an infinite split): a standalone marker run is
						// removed whole; a run that also carried post-break content
						// keeps the content but loses the marker element.
						const lead = newParagraph.children[0];
						if (lead && this.isPageBreakRun(lead)) {
							if ((lead.children?.length ?? 0) <= 1) {
								newParagraph.children = newParagraph.children.slice(1);
							} else {
								const bi = lead.children.findIndex(this.isPageBreakElement.bind(this));
								if (bi >= 0) {
									// again a clone - the run belongs to the model
									newParagraph.children = ([
										{ ...lead, children: lead.children.filter((c, k) => k !== bi) }
									] as OpenXmlElement[]).concat(newParagraph.children.slice(1));
								}
							}
						}
						if (newParagraph.children.length)
							list.splice(ei + 1, 0, newParagraph);
					}
				}
			}
		}

		let currentSectProps = null;

		for (let i = result.length - 1; i >= 0; i--) {
			if (result[i].sectProps == null) {
				result[i].sectProps = currentSectProps ?? defaultProps;
			} else {
				currentSectProps = result[i].sectProps
			}
		}

		return result;
	}

	groupByPageBreaks(sections: Section[]): Section[][] {
		let current = [];
		let prev: SectionProperties;
		const result: Section[][] = [current];

		for (let s of sections) {
			current.push(s);

			if (this.options.ignoreLastRenderedPageBreak || s.pageBreak || this.isPageBreakSection(prev, s.sectProps))
				result.push(current = []);

			prev = s.sectProps;
		}

		return result.filter(x => x.length > 0);
	}

	renderWrapper(children: HTMLElement[]) {
		return this.h({ tagName: "div", className: `${this.className}-wrapper`, children });
	}

	renderDefaultStyle() {
		var c = this.className;
		var wrapperStyle = `
.${c}-wrapper { background: gray; padding: 30px; padding-bottom: 0px; display: flex; flex-flow: column; align-items: center; } 
.${c}-wrapper>section.${c} { background: white; box-shadow: 0 0 10px rgba(0, 0, 0, 0.5); margin-bottom: 30px; }`;
		if (this.options.hideWrapperOnPrint) {
			wrapperStyle = `@media not print { ${wrapperStyle} }`;
		}
		var styleText = `${wrapperStyle}
.${c} { color: black; hyphens: auto; text-underline-position: from-font; text-rendering: geometricPrecision; }
section.${c} { box-sizing: border-box; display: flex; flex-flow: column nowrap; position: relative; overflow: hidden; }
section.${c}>article { margin-bottom: auto; z-index: 1; }
section.${c}>header { z-index: 0; }
section.${c}>footer { z-index: 0; }
.${c} table { border-collapse: collapse; }
.${c} table td, .${c} table th { vertical-align: top; }
.${c} p { margin: 0pt; min-height: 1em; }
.${c} span { white-space: pre-wrap; overflow-wrap: break-word; }
.${c} a { color: inherit; text-decoration: inherit; }
.${c} svg { fill: transparent; }
`;

		if (this.options.renderComments) {
			styleText += `
.${c}-comment-ref { cursor: default; }
.${c}-comment-popover { display: none; z-index: 1000; padding: 0.5rem; background: white; position: absolute; box-shadow: 0 0 0.25rem rgba(0, 0, 0, 0.25); width: 30ch; }
.${c}-comment-ref:hover~.${c}-comment-popover { display: block; }
.${c}-comment-author,.${c}-comment-date { font-size: 0.875rem; color: #888; }
`
		};

		return [
			this.h({ tagName: "#comment", children: ["docxjs library predefined styles"] }),
			this.h({ tagName: "style", children: [styleText] })
		];
	}

	// renderNumbering2(numberingPart: NumberingPartProperties, container: HTMLElement): HTMLElement {
	//     let css = "";
	//     const numberingMap = keyBy(numberingPart.abstractNumberings, x => x.id);
	//     const bulletMap = keyBy(numberingPart.bulletPictures, x => x.id);
	//     const topCounters = [];

	//     for(let num of numberingPart.numberings) {
	//         const absNum = numberingMap[num.abstractId];

	//         for(let lvl of absNum.levels) {
	//             const className = this.numberingClass(num.id, lvl.level);
	//             let listStyleType = "none";

	//             if(lvl.text && lvl.format == 'decimal') {
	//                 const counter = this.numberingCounter(num.id, lvl.level);

	//                 if (lvl.level > 0) {
	//                     css += this.styleToString(`p.${this.numberingClass(num.id, lvl.level - 1)}`, {
	//                         "counter-reset": counter
	//                     });
	//                 } else {
	//                     topCounters.push(counter);
	//                 }

	//                 css += this.styleToString(`p.${className}:before`, {
	//                     "content": this.levelTextToContent(lvl.text, num.id),
	//                     "counter-increment": counter
	//                 });
	//             } else if(lvl.bulletPictureId) {
	//                 let pict = bulletMap[lvl.bulletPictureId];
	//                 let variable = `--${this.className}-${pict.referenceId}`.toLowerCase();

	//                 css += this.styleToString(`p.${className}:before`, {
	//                     "content": "' '",
	//                     "display": "inline-block",
	//                     "background": `var(${variable})`
	//                 }, pict.style);

	//                 this.document.loadNumberingImage(pict.referenceId).then(data => {
	//                     var text = `.${this.className}-wrapper { ${variable}: url(${data}) }`;
	//                     container.appendChild(createStyleElement(text));
	//                 });
	//             } else {
	//                 listStyleType = this.numFormatToCssValue(lvl.format);
	//             }

	//             css += this.styleToString(`p.${className}`, {
	//                 "display": "list-item",
	//                 "list-style-position": "inside",
	//                 "list-style-type": listStyleType,
	//                 //TODO
	//                 //...num.style
	//             });
	//         }
	//     }

	//     if (topCounters.length > 0) {
	//         css += this.styleToString(`.${this.className}-wrapper`, {
	//             "counter-reset": topCounters.join(" ")
	//         });
	//     }

	//     return createStyleElement(css);
	// }

	async renderNumbering(numberings: IDomNumbering[]) {
		var styleText = "";
		var resetCounters = [];

		for (var num of numberings) {
			var selector = `p.${this.numberingClass(num.id, num.level)}`;
			var listStyleType = "none";

			if (num.bullet) {
				let valiable = `--${this.className}-${num.bullet.src}`.toLowerCase();

				styleText += this.styleToString(`${selector}:before`, {
					"content": "' '",
					"display": "inline-block",
					"background": `var(${valiable})`
				}, num.bullet.style);

				try {
					const imgData = await this.document.loadNumberingImage(num.bullet.src);
					styleText += `${this.rootSelector} { ${valiable}: url(${imgData}) }`;
				} catch(e) {
					if (this.options.debug) console.warn(`Can't load numbering image with src ${num.bullet.src}`);
				}
			}
			else if (num.levelText) {
				let counter = this.numberingCounter(num.id, num.level);
				const counterReset = counter + " " + (num.start - 1);
				if (num.level > 0) {
					styleText += this.styleToString(`p.${this.numberingClass(num.id, num.level - 1)}`, {
						"counter-set": counterReset
					});
				}
				// reset all level counters with start value
				resetCounters.push(counterReset);

				// the number's run style: the level's own rPr wins, but fall
				// back to the linked paragraph style's run properties so the
				// number matches the heading's weight/font (Word behavior -
				// otherwise the ::before only inherits the paragraph element,
				// which carries no run formatting)
				var numRunStyle = num.rStyle;
				var linkedStyle = this.findStyle(num.pStyleName);
				var linkedRun = linkedStyle?.styles?.find(s => s.target == "span")?.values;
				if (linkedRun)
					numRunStyle = { ...linkedRun, ...num.rStyle };

				var beforeValues: Record<string, string> = {
					"content": this.levelTextToContent(num.levelText, num.suff, num.id, this.numFormatToCssValue(num.format)),
					"counter-increment": counter,
					...numRunStyle,
				};

				// a "tab" suffix should advance to a tab stop, but a CSS tab
				// character in generated content collapses to a thin space.
				// Emulate it: make the marker an inline-block up to the tab stop,
				// so the text starts there. With a hanging indent Word's implicit
				// tab stop sits at the paragraph's left indent - the marker box
				// spans exactly the hanging width (= negative text-indent);
				// otherwise fall back to one default tab stop.
				if (num.suff == null || num.suff == "tab") {
					var hangingIndent = num.pStyle?.["text-indent"];
					beforeValues["display"] = "inline-block";
					// levels declaring hanging="0" parse to "-0.00pt" - only a
					// real (negative) hanging indent replaces the default stop
					beforeValues["min-width"] = (hangingIndent && parseFloat(hangingIndent) < 0)
						? hangingIndent.substring(1)
						: (this.document.settingsPart?.settings?.defaultTabStop ?? "36pt");
					// baseline-aligned inline-blocks inflate the line box by
					// their descent; top keeps the marker's internal baseline
					// on the text baseline (same font/line-height) without that
					beforeValues["vertical-align"] = "top";
					// an inline-block is a block container: the paragraph's
					// (negative) text-indent inherits into it and would indent
					// the marker box's OWN first line again, painting the number
					// a full hanging width left of the box (and of the column)
					beforeValues["text-indent"] = "0";
				}

				styleText += this.styleToString(`${selector}:before`, beforeValues);
			}
			else {
				listStyleType = this.numFormatToCssValue(num.format);
			}

			styleText += this.styleToString(selector, {
				"display": "list-item",
				"list-style-position": "inside",
				"list-style-type": listStyleType,
				...num.pStyle
			});
		}

		if (resetCounters.length > 0) {
			styleText += this.styleToString(this.rootSelector, {
				"counter-reset": resetCounters.join(" ")
			});
		}

		return [
			this.h({ tagName: "#comment", children: ["docxjs document numbering styles"] }),
			this.h({ tagName: "style", children: [styleText] })
		];
	}

	renderStyles(styles: IDomStyle[]) {
		var styleText = "";
		const stylesMap = this.styleMap;
		const defautStyles = keyBy(styles.filter(s => s.isDefault), s => s.target);

		for (const style of styles) {
			var subStyles = style.styles;

			if (style.linked) {
				var linkedStyle = style.linked && stylesMap[style.linked];

				if (linkedStyle)
					subStyles = subStyles.concat(linkedStyle.styles);
				else if (this.options.debug)
					console.warn(`Can't find linked style ${style.linked}`);
			}

			for (const subStyle of subStyles) {
				//TODO temporary disable modificators until test it well
				var selector = `${style.target ?? ''}.${style.cssName}`; //${subStyle.mod ?? ''} 

				if (style.target != subStyle.target)
					selector += ` ${subStyle.target}`;

				if (defautStyles[style.target] == style)
					selector = `.${this.className} ${style.target}, ` + selector;

				styleText += this.styleToString(selector, subStyle.values);
			}
		}

		return [
			this.h({ tagName: "#comment", children: ["docxjs document styles"] }),
			this.h({ tagName: "style", children: [styleText] })
		];
	}

	renderNotes(noteIds: string[], notesMap: Record<string, WmlBaseNote>) {
		var notes = noteIds.map(id => notesMap[id]).filter(x => x);

		if (notes.length > 0) {
			return this.h({ tagName: "ol", children: this.renderElements(notes) });
		}
	}

	renderElement(elem: OpenXmlElement): Node | Node[] {
		switch (elem.type) {
			case DomType.Paragraph:
				return this.renderParagraph(elem as WmlParagraph);

			case DomType.BookmarkStart:
				return this.renderBookmarkStart(elem as WmlBookmarkStart);

			case DomType.BookmarkEnd:
				return null; //ignore bookmark end

			case DomType.Run:
				return this.renderRun(elem as WmlRun);

			case DomType.SimpleField:
				return this.renderSimpleField(elem as WmlFieldSimple);

			case DomType.Table:
				return this.renderTable(elem);

			case DomType.Row:
				return this.renderTableRow(elem);

			case DomType.Cell:
				return this.renderTableCell(elem);

			case DomType.Hyperlink:
				return this.renderHyperlink(elem);
			
			case DomType.SmartTag:
				return this.renderSmartTag(elem);

			case DomType.Drawing:
				return this.renderDrawing(elem);

			case DomType.Image:
				return this.renderImage(elem as IDomImage);

			case DomType.Text:
				return this.renderText(elem as WmlText);

			case DomType.Text:
				return this.renderText(elem as WmlText);

			case DomType.DeletedText:
				return this.renderDeletedText(elem as WmlText);
	
			case DomType.Tab:
				return this.renderTab(elem);

			case DomType.Symbol:
				return this.renderSymbol(elem as WmlSymbol);

			case DomType.Break:
				return this.renderBreak(elem as WmlBreak);

			case DomType.Footer:
				return this.renderContainer(elem, "footer");

			case DomType.Header:
				return this.renderContainer(elem, "header");

			case DomType.Footnote:
			case DomType.Endnote:
				return this.renderContainer(elem, "li");

			case DomType.FootnoteReference:
				return this.renderFootnoteReference(elem as WmlNoteReference);

			case DomType.EndnoteReference:
				return this.renderEndnoteReference(elem as WmlNoteReference);

			case DomType.NoBreakHyphen:
				return this.h({ tagName: "wbr" });

			case DomType.VmlPicture:
				return this.renderVmlPicture(elem);

			case DomType.VmlElement:
				return this.renderVmlElement(elem as VmlElement);
	
			case DomType.MmlMath:
				return this.renderContainerNS(elem, ns.mathML, "math", { xmlns: ns.mathML });
	
			case DomType.MmlMathParagraph:
				return this.renderContainer(elem, "span");

			case DomType.MmlFraction:
				return this.renderContainerNS(elem, ns.mathML, "mfrac");

			case DomType.MmlBase:
				return this.renderContainerNS(elem, ns.mathML, 
					elem.parent.type == DomType.MmlMatrixRow ? "mtd" : "mrow");

			case DomType.MmlNumerator:
			case DomType.MmlDenominator:
			case DomType.MmlFunction:
			case DomType.MmlLimit:
			case DomType.MmlBox:
				return this.renderContainerNS(elem, ns.mathML, "mrow");

			case DomType.MmlGroupChar:
				return this.renderMmlGroupChar(elem);

			case DomType.MmlLimitLower:
				return this.renderContainerNS(elem, ns.mathML, "munder");

			case DomType.MmlMatrix:
				return this.renderContainerNS(elem, ns.mathML, "mtable");

			case DomType.MmlMatrixRow:
				return this.renderContainerNS(elem, ns.mathML, "mtr");
	
			case DomType.MmlRadical:
				return this.renderMmlRadical(elem);

			case DomType.MmlSuperscript:
				return this.renderContainerNS(elem, ns.mathML, "msup");

			case DomType.MmlSubscript:
				return this.renderContainerNS(elem, ns.mathML, "msub");

			case DomType.MmlDegree:
			case DomType.MmlSuperArgument:
			case DomType.MmlSubArgument:
				return this.renderContainerNS(elem, ns.mathML, "mn");

			case DomType.MmlFunctionName:
				return this.renderContainerNS(elem, ns.mathML, "ms");
	
			case DomType.MmlDelimiter:
				return this.renderMmlDelimiter(elem);

			case DomType.MmlRun:
				return this.renderMmlRun(elem);

			case DomType.MmlNary:
				return this.renderMmlNary(elem);

			case DomType.MmlPreSubSuper:
				return this.renderMmlPreSubSuper(elem);

			case DomType.MmlBar:
				return this.renderMmlBar(elem);
	
			case DomType.MmlEquationArray:
				return this.renderMllList(elem);

			case DomType.Inserted:
				return this.renderInserted(elem as WmlChange);

			case DomType.Deleted:
				return this.renderDeleted(elem as WmlChange);

			case DomType.CommentRangeStart:
				return this.renderCommentRangeStart(elem);

			case DomType.CommentRangeEnd:
				return this.renderCommentRangeEnd(elem);

			case DomType.CommentReference:
				return this.renderCommentReference(elem);

			case DomType.AltChunk:
				return this.renderAltChunk(elem);
		}

		return null;
	}
	renderElements(elems: OpenXmlElement[], into?: Node): Node[] {
		if (elems == null)
			return null;

		var result = elems.flatMap(e => this.renderElement(e)).filter(e => e != null);

		if (into)
			result.forEach(c => into.appendChild(isString(c) ? document.createTextNode(c) : c));

		return result;
	}

	renderContainer<T extends keyof HTMLElementTagNameMap>(elem: OpenXmlElement, tagName: T, props?: Record<string, any>): HTMLElementTagNameMap[T] {
		return this.h({ tagName, children: this.renderElements(elem.children), ...props }) as any;
	}

	renderContainerNS(elem: OpenXmlElement, ns: ns, tagName: string, props?: Record<string, any>) {
		return this.h({ ns, tagName, children: this.renderElements(elem.children), ...props });
	}

	renderParagraph(elem: WmlParagraph) {
		var result = this.toHTML(elem, ns.html, "p");

		const style = this.findStyle(elem.styleName);
		elem.tabs ??= style?.paragraphProps?.tabs;  //TODO

		// a continuation of a paragraph split at a measured page break must not
		// repeat its number nor re-increment the numbering counter
		const numbering = (this.options.measuredPagination && (elem as any).$splitCont)
			? null
			: (elem.numbering ?? style?.paragraphProps?.numbering);

		if (numbering) {
			result.classList.add(this.numberingClass(numbering.id, numbering.level));
		}

		if (this.options.measuredPagination && (elem as any).$mi != null)
			result.setAttribute("data-docx-mi", String((elem as any).$mi));

		return result;
	}

	renderHyperlink(elem: WmlHyperlink) {
		const res = this.toH(elem, ns.html, "a");
		res.href = '';

		if (elem.id) {
			const rel = this.document.documentPart.rels.find(it => it.id == elem.id && it.targetMode === "External");
			res.href = rel?.target ?? res.href;
		}

		if (elem.anchor) {
			res.href += `#${elem.anchor}`;
		}

		return this.h(res);
	}
	
	renderSmartTag(elem: WmlSmartTag) {
		return this.renderContainer(elem, "span");
	}
	
	renderCommentRangeStart(commentStart: WmlCommentRangeStart) {
		if (!this.options.renderComments)
			return null;

		const rng = new Range();
		this.commentHighlight?.add(rng);

		const result = this.h({ tagName: "#comment", children: [`start of comment #${commentStart.id}`] });
		this.later(() => rng.setStart(result, 0));
		this.commentMap[commentStart.id] = rng;

		return result
	}

	renderCommentRangeEnd(commentEnd: WmlCommentRangeStart) {
		if (!this.options.renderComments)
			return null;

		const rng = this.commentMap[commentEnd.id];
		const result = this.h({ tagName: "#comment", children: [`end of comment #${commentEnd.id}`] });
		this.later(() => rng?.setEnd(result, 0));

		return result;
	}

	renderCommentReference(commentRef: WmlCommentReference) {
		if (!this.options.renderComments)
			return null;

		var comment = this.document.commentsPart?.commentMap[commentRef.id];

		if (!comment)
			return null;

		const commentRefEl = this.h({ tagName: "span", className: `${this.className}-comment-ref`, children: ['💬'] });
		const commentsContainerEl = this.h({
			tagName: "div", className: `${this.className}-comment-popover`, children: [
				this.h({ tagName: 'div', className: `${this.className}-comment-author`, children: [comment.author] }),
				this.h({ tagName: 'div', className: `${this.className}-comment-date`, children: [new Date(comment.date).toLocaleString()] }),
				...this.renderElements(comment.children)
			]
		});

		return this.h({ tagName:  "#fragment", children: [
			this.h({ tagName: "#comment", children: [`comment #${comment.id} by ${comment.author} on ${comment.date}`] }),
			commentRefEl,
			commentsContainerEl
		] });
	}

	renderAltChunk(elem: WmlAltChunk) {
		if (!this.options.renderAltChunks)
			return null;

		var result = this.h({ tagName: "iframe" }) as HTMLIFrameElement;
		
		this.tasks.push(this.document.loadAltChunk(elem.id, this.currentPart).then(x => {
			result.srcdoc = x;
		}));

		return result;
	}

	renderDrawing(elem: OpenXmlElement) {
		var result = this.toHTML(elem, ns.html, "div");

		result.style.display = "inline-block";
		result.style.position = "relative";
		result.style.textIndent = "0px";

		// Expose the anchor's reference frames as attributes (CSS custom properties
		// are dropped when cssStyle is applied to the element) so a post-attach
		// pass can move the drawing from the inline position it was laid out at to
		// the base its offset is measured from - the page, the margins or the
		// anchoring paragraph - once the page geometry is known.
		const relH = elem.cssStyle?.["--docx-anchor-rel-h"];
		if (relH)
			result.setAttribute("data-docx-anchor-h", relH);

		const relV = elem.cssStyle?.["--docx-anchor-rel-v"];
		if (relV)
			result.setAttribute("data-docx-anchor-v", relV);

		return result;
	}

	renderImage(elem: IDomImage) {
		let result = this.toHTML(elem, ns.html, "img", []);
		let transform = elem.cssStyle?.transform;

		if (elem.srcRect && elem.srcRect.some(x => x != 0)) {
			var [left, top, right, bottom] = elem.srcRect;
			transform = `scale(${1 / (1 - left - right)}, ${1 / (1 - top - bottom)})`;
			result.style['clip-path'] = `rect(${(100 * top).toFixed(2)}% ${(100 * (1 - right)).toFixed(2)}% ${(100 * (1 - bottom)).toFixed(2)}% ${(100 * left).toFixed(2)}%)`;
		}

		if (elem.rotation)
			transform = `rotate(${elem.rotation}deg) ${transform ?? ''}`;

		result.style.transform = transform?.trim();

		if (this.document) {
			this.tasks.push(this.document.loadDocumentImage(elem.src, this.currentPart).then(x => {
				result.src = x;
			}));
		}

		return result;
	}

	renderText(elem: WmlText) {
		return this.h(elem.text);
	}

	renderDeletedText(elem: WmlText) {
		return this.options.renderChanges ? this.renderText(elem) : null;
	}

	renderBreak(elem: WmlBreak) {
		return elem.break == "textWrapping" ? this.h({ tagName: "br" }) : null;
	}

	renderInserted(elem: WmlChange): Node | Node[] {
		if (this.options.renderChanges)
			return this.renderChange(elem, "ins");

		return this.renderElements(elem.children);
	}

	renderDeleted(elem: WmlChange): Node {
		if (this.options.renderChanges)
			return this.renderChange(elem, "del");

		return null;
	}

	renderChange(elem: WmlChange, tag: string) {
		return this.renderContainer(elem, tag as any, {
			dateTime: elem.date		
		});
	}

	renderSymbol(elem: WmlSymbol) {
		return this.h({ tagName: "span", children: [String.fromCharCode(elem.char)], style: { fontFamily: elem.font } });
	}

	renderFootnoteReference(elem: WmlNoteReference) {
		this.currentFootnoteIds.push(elem.id);
		return this.h({ tagName: "sup", children: [`${this.currentFootnoteIds.length}`] });
	}

	renderEndnoteReference(elem: WmlNoteReference) {
		this.currentEndnoteIds.push(elem.id);
		return this.h({ tagName: "sup", children: [`${this.currentEndnoteIds.length}`] });
	}

	renderTab(elem: OpenXmlElement) {
		var tabSpan = this.h({ tagName: "span", children: ["\u2003"] }) as HTMLElement;//"&nbsp;";

		if (this.options.experimental) {
			tabSpan.className = this.tabStopClass();
			var stops = findParent<WmlParagraph>(elem, DomType.Paragraph)?.tabs;
			this.currentTabs.push({ stops, span: tabSpan });
		}

		return tabSpan;
	}

	renderBookmarkStart(elem: WmlBookmarkStart) {
		return this.h({ tagName: "span", id: elem.name });
	}

	renderRun(elem: WmlRun) {
		if (elem.fieldRun || elem.fieldCodeContent)
			return null;

		let children = this.renderElements(elem.children);

		if (elem.verticalAlign) {
			children = [this.h({ tagName: elem.verticalAlign, children: this.renderElements(elem.children) })];
		}

		const result = this.toHTML(elem, ns.html, "span", children);

		if (elem.id)
			result.id = elem.id;

		// mark PAGE/NUMPAGES/SECTIONPAGES result runs so measured pagination can
		// fill the live per-page value (see fillPageFields)
		if (elem.fieldValue && this.options.measuredPagination)
			result.setAttribute("data-docx-field", elem.fieldValue);

		return result;
	}

	// A simple field (w:fldSimple). Only PAGE/NUMPAGES/SECTIONPAGES fields are
	// rendered, and only under measuredPagination (where the live value can be
	// computed) - matching the default renderer, which does not render simple
	// fields at all. The cached result is rendered as a placeholder and then
	// overwritten per page by fillPageFields.
	renderSimpleField(elem: WmlFieldSimple) {
		if (!elem.fieldValue || !this.options.measuredPagination)
			return null;

		const result = this.toHTML(elem, ns.html, "span", this.renderElements(elem.children));
		result.setAttribute("data-docx-field", elem.fieldValue);
		return result;
	}

	// Fills PAGE/NUMPAGES/SECTIONPAGES field placeholders with live values now
	// that the final pages are known. SECTIONPAGES counts pages sharing a
	// section (data-docx-sect, stamped per page in measured mode).
	fillPageFields(pages: HTMLElement[]) {
		const sectKey = (s: HTMLElement, i: number) => s.getAttribute("data-docx-sect") ?? ("solo" + i);
		const perSect: Record<string, number> = {};
		pages.forEach((s, i) => { const k = sectKey(s, i); perSect[k] = (perSect[k] ?? 0) + 1; });
		pages.forEach((s, i) => {
			const vals: Record<string, string> = {
				PAGE: String(i + 1),
				NUMPAGES: String(pages.length),
				SECTIONPAGES: String(perSect[sectKey(s, i)])
			};
			s.querySelectorAll<HTMLElement>("[data-docx-field]").forEach(el => {
				const v = vals[el.getAttribute("data-docx-field")];
				if (v != null) el.textContent = v;
			});
		});
	}

	renderTable(elem: WmlTable) {
		this.tableCellPositions.push(this.currentCellPosition);
		this.tableVerticalMerges.push(this.currentVerticalMerge);
		this.currentVerticalMerge = {};
		this.currentCellPosition = { col: 0, row: 0 };

		const children = [];

		if (elem.columns)
			children.push(this.renderTableColumns(elem.columns));

		children.push(...this.renderElements(elem.children));

		this.currentVerticalMerge = this.tableVerticalMerges.pop();
		this.currentCellPosition = this.tableCellPositions.pop();
		const result = this.toHTML(elem, ns.html, "table", children) as HTMLElement;

		if (this.options.measuredPagination && (elem as any).$mi != null)
			result.setAttribute("data-docx-mi", String((elem as any).$mi));

		// a segment of a table split at its own row breaks: tell the measured pass
		// which row of the original table this segment starts at
		if (this.options.measuredPagination && (elem as any).$rowbase)
			result.setAttribute("data-docx-rowbase", String((elem as any).$rowbase));

		return result;
	}

	renderTableColumns(columns: WmlTableColumn[]) {
		const children = columns.map(x => this.h({ tagName: "col", style: { width: x.width } }));
		return this.h({ tagName: "colgroup", children });
	}

	renderTableRow(elem: WmlTableRow) {
		this.currentCellPosition.col = 0;

		const children = [];

		if (elem.gridBefore)
			children.push(this.renderTableCellPlaceholder(elem.gridBefore));

		children.push(...this.renderElements(elem.children));

		if (elem.gridAfter)
			children.push(this.renderTableCellPlaceholder(elem.gridAfter));

		this.currentCellPosition.row++;

		return this.toHTML(elem, ns.html, "tr", children) as HTMLTableRowElement;
	}

	renderTableCellPlaceholder(colSpan: number) {
		return this.h({ tagName: "td", colSpan, style: { border: "none" } });
	}

	renderTableCell(elem: WmlTableCell) {
		let result = this.toHTML(elem, ns.html, "td");

		const key = this.currentCellPosition.col;

		if (elem.verticalMerge) {
			if (elem.verticalMerge == "restart") {
				this.currentVerticalMerge[key] = result;
				result.rowSpan = 1;
			} else if (this.currentVerticalMerge[key]) {
				this.currentVerticalMerge[key].rowSpan += 1;
				result.style.display = "none";
			}
		} else {
			this.currentVerticalMerge[key] = null;
		}

		if (elem.span)
			result.colSpan = elem.span;

		this.currentCellPosition.col += result.colSpan;

		return result;
	}

	renderVmlPicture(elem: OpenXmlElement) {
		return this.renderContainer(elem, "div");
	}

	renderVmlElement(elem: VmlElement): SVGElement {
		var container = this.h({ ns: ns.svg, tagName: "svg", style: elem.cssStyleText }) as SVGElement;

		const result = this.renderVmlChildElement(elem);

		if (elem.imageHref?.id) {
			this.tasks.push(this.document?.loadDocumentImage(elem.imageHref.id, this.currentPart)
				.then(x => result.setAttribute("href", x)));
		}

		container.appendChild(result);

		requestAnimationFrame(() => {
			const bb = (container.firstElementChild as any).getBBox();

			container.setAttribute("width", `${Math.ceil(bb.x +  bb.width)}`);
			container.setAttribute("height", `${Math.ceil(bb.y + bb.height)}`);
		});

		return container;
	}

	renderVmlChildElement(elem: VmlElement): any {
		const result = this.createSvgElement(elem.tagName as any);
		Object.entries(elem.attrs).forEach(([k, v]) => result.setAttribute(k, v));

		for (let child of elem.children) {
			if (child.type == DomType.VmlElement) {
				result.appendChild(this.renderVmlChildElement(child as VmlElement));
			} else {
				result.appendChild(...asArray(this.renderElement(child as any)));
			}
		}

		return result;
	}

	renderMmlRadical(elem: OpenXmlElement) {
		const base = elem.children.find(el => el.type == DomType.MmlBase);

		if (elem.props?.hideDegree) {
			return this.createMathMLElement("msqrt", null, this.renderElements([base]));
		}

		const degree = elem.children.find(el => el.type == DomType.MmlDegree);
		return this.createMathMLElement("mroot", null, this.renderElements([base, degree]));
	}

	renderMmlDelimiter(elem: OpenXmlElement) {		
		const children = [];

		children.push(this.createMathMLElement("mo", null, [elem.props.beginChar ?? '(']));
		children.push(...this.renderElements(elem.children));
		children.push(this.createMathMLElement("mo", null, [elem.props.endChar ?? ')']));

		return this.createMathMLElement("mrow", null, children);
	}

	renderMmlNary(elem: OpenXmlElement) {		
		const children = [];
		const grouped = keyBy(elem.children, x => x.type);

		const sup = grouped[DomType.MmlSuperArgument];
		const sub = grouped[DomType.MmlSubArgument];
		const supElem = sup ? this.createMathMLElement("mo", null, asArray(this.renderElement(sup))) : null;
		const subElem = sub ? this.createMathMLElement("mo", null, asArray(this.renderElement(sub))) : null;

		const charElem = this.createMathMLElement("mo", null, [elem.props?.char ?? '\u222B']);

		if (supElem || subElem) {
			children.push(this.createMathMLElement("munderover", null, [charElem, subElem, supElem]));
		} else if(supElem) {
			children.push(this.createMathMLElement("mover", null, [charElem, supElem]));
		} else if(subElem) {
			children.push(this.createMathMLElement("munder", null, [charElem, subElem]));
		} else {
			children.push(charElem);
		}

		children.push(...this.renderElements(grouped[DomType.MmlBase].children));

		return this.createMathMLElement("mrow", null, children);
	}

	renderMmlPreSubSuper(elem: OpenXmlElement) {
		const children = [];
		const grouped = keyBy(elem.children, x => x.type);

		const sup = grouped[DomType.MmlSuperArgument];
		const sub = grouped[DomType.MmlSubArgument];
		const supElem = sup ? this.createMathMLElement("mo", null, asArray(this.renderElement(sup))) : null;
		const subElem = sub ? this.createMathMLElement("mo", null, asArray(this.renderElement(sub))) : null;
		const stubElem = this.createMathMLElement("mo", null);

		children.push(this.createMathMLElement("msubsup", null, [stubElem, subElem, supElem]));
		children.push(...this.renderElements(grouped[DomType.MmlBase].children));

		return this.createMathMLElement("mrow", null, children);
	}

	renderMmlGroupChar(elem: OpenXmlElement) {
		const tagName = elem.props.verticalJustification === "bot" ? "mover" : "munder";
		const result = this.renderContainerNS(elem, ns.mathML, tagName);

		if (elem.props.char) {
			result.appendChild(this.createMathMLElement("mo", null, [elem.props.char]));
		}

		return result;
	}

	renderMmlBar(elem: OpenXmlElement) {
		const style = {} as any;

		switch(elem.props.position) {
			case "top": style.textDecoration = "overline"; break
			case "bottom": style.textDecoration = "underline"; break
		}

		return this.renderContainerNS(elem, ns.mathML, "mrow", { style }) as MathMLElement;
	}

	renderMmlRun(elem: OpenXmlElement) {
		return this.toHTML(elem, ns.mathML, "ms");
	}

	renderMllList(elem: OpenXmlElement) {
		const children = this.renderElements(elem.children).map(x => this.createMathMLElement("mtr", null, [
			this.createMathMLElement("mtd", null, [x])
		]));

		return this.toHTML(elem, ns.mathML, "mtable", children);
	}

	toH(elem: OpenXmlElement, ns: ns, tagName: string, children: Node[] = null) {
		const { "$lang": lang, ...style } = elem.cssStyle ?? {};
		const className = cx(elem.className, elem.styleName && this.processStyleName(elem.styleName));
		return { ns, tagName, className, lang, style, children: children ?? this.renderElements(elem.children) } as any;
	}

	toHTML(elem: OpenXmlElement, ns: ns, tagName: string, children: Node[] = null) {
		return this.h(this.toH(elem, ns, tagName, children)) as any;
	}

	findStyle(styleName: string) {
		return styleName && this.styleMap?.[styleName];
	}

	numberingClass(id: string, lvl: number) {
		return `${this.className}-num-${id}-${lvl}`;
	}

	tabStopClass() {
		return `${this.className}-tab-stop`;
	}

	styleToString(selectors: string, values: Record<string, string>, cssText: string = null) {
		let result = `${selectors} {\r\n`;

		for (const key in values) {
			if (key.startsWith('$'))
				continue;
			
			result += `  ${key}: ${values[key]};\r\n`;
		}

		if (cssText)
			result += cssText;

		return result + "}\r\n";
	}

	numberingCounter(id: string, lvl: number) {
		return `${this.className}-num-${id}-${lvl}`;
	}

	levelTextToContent(text: string, suff: string, id: string, numformat: string) {
		const suffMap = {
			"tab": "\\9",
			"space": "\\a0",
		};

		var result = text.replace(/%\d*/g, s => {
			let lvl = parseInt(s.substring(1), 10) - 1;
			return `"counter(${this.numberingCounter(id, lvl)}, ${numformat})"`;
		});

		return `"${result}${suffMap[suff] ?? ""}"`;
	}

	numFormatToCssValue(format: string) {
		var mapping = {
			none: "none",
			bullet: "disc",
			decimal: "decimal",
			lowerLetter: "lower-alpha",
			upperLetter: "upper-alpha",
			lowerRoman: "lower-roman",
			upperRoman: "upper-roman",
			decimalZero: "decimal-leading-zero", // 01,02,03,...
			// ordinal: "", // 1st, 2nd, 3rd,...
			// ordinalText: "", //First, Second, Third, ...
			// cardinalText: "", //One,Two Three,...
			// numberInDash: "", //-1-,-2-,-3-, ...
			// hex: "upper-hexadecimal",
			aiueo: "katakana",
			aiueoFullWidth: "katakana",
			chineseCounting: "simp-chinese-informal",
			chineseCountingThousand: "simp-chinese-informal",
			chineseLegalSimplified: "simp-chinese-formal", // 中文大写
			chosung: "hangul-consonant",
			ideographDigital: "cjk-ideographic",
			ideographTraditional: "cjk-heavenly-stem", // 十天干
			ideographLegalTraditional: "trad-chinese-formal",
			ideographZodiac: "cjk-earthly-branch", // 十二地支
			iroha: "katakana-iroha",
			irohaFullWidth: "katakana-iroha",
			japaneseCounting: "japanese-informal",
			japaneseDigitalTenThousand: "cjk-decimal",
			japaneseLegal: "japanese-formal",
			thaiNumbers: "thai",
			koreanCounting: "korean-hangul-formal",
			koreanDigital: "korean-hangul-formal",
			koreanDigital2: "korean-hanja-informal",
			hebrew1: "hebrew",
			hebrew2: "hebrew",
			hindiNumbers: "devanagari",
			ganada: "hangul",
			taiwaneseCounting: "cjk-ideographic",
			taiwaneseCountingThousand: "cjk-ideographic",
			taiwaneseDigital:  "cjk-decimal",
		};

		return mapping[format] ?? format;
	}

	refreshTabStops() {
		if (!this.options.experimental)
			return;

		setTimeout(() => {
			const pixelToPoint = computePixelToPoint();

			for (let tab of this.currentTabs) {
				updateTabStop(tab.span, tab.stops, this.defaultTabSize, pixelToPoint);
			}
		}, 500);
	}

	createElementNS(ns: any, tagName: string, props?: Partial<Record<any, any>>, children?: any[]) {
		return this.h({ ns, tagName, children, ...props }) as any;
	}

	createElement<T extends keyof HTMLElementTagNameMap>(tagName: T, props?: Partial<Record<keyof HTMLElementTagNameMap[T], any>>, children?: any[]): HTMLElementTagNameMap[T] {
		return this.createElementNS(ns.html, tagName, props, children);
	}

	createMathMLElement<T extends keyof MathMLElementTagNameMap>(tagName: T, props?: Partial<Record<keyof MathMLElementTagNameMap[T], any>>, children?: any[]): MathMLElementTagNameMap[T] {
		return this.createElementNS(ns.mathML, tagName, props, children);
	}

	createSvgElement<T extends keyof SVGElementTagNameMap>(tagName: T, props?: Partial<Record<keyof SVGElementTagNameMap[T], any>>, children?: any[]): SVGElementTagNameMap[T] {
		return this.createElementNS(ns.svg, tagName, props, children);
	}

	later(func: Function) { 
		this.postRenderTasks.push(func);
	}
}

function findParent<T extends OpenXmlElement>(elem: OpenXmlElement, type: DomType): T {
	var parent = elem.parent;

	while (parent != null && parent.type != type)
		parent = parent.parent;

	return <T>parent;
}

