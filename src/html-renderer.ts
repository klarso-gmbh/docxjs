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
import { WmlParagraph, ParagraphTab } from './document/paragraph';
import { asArray, encloseFontFamily, escapeClassName, isString, keyBy, mergeDeep } from './utils';
import { computePixelToPoint, updateTabStop } from './javascript';
import { FontTablePart } from './font-table/font-table';
import { FooterHeaderReference, SectionProperties, SectionType } from './document/section';
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
	resolveBreaksToDocument, hyphenateModel, getAutoLineMul, removeSavedPageBreaks
} from './measured-pagination';
import { tabStopsAttr, resolveTabStops, defaultTabPoints } from './tab-stops';
import { addUpParagraphSpacing, PREV_AFTER_RESET } from './paragraph-spacing';

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
	// the default paragraph style (w:default="1"), which formats every
	// paragraph without a w:pStyle
	defaultParagraphStyleId: string = null;
	// numberingAsText: the levels of every w:num by level, and the running
	// count of each (reset per render pass)
	numberingLevels: Record<string, Record<number, IDomNumbering>> = {};
	numberingCounters: Record<string, number[]> = {};
	currentPart: Part = null;

	tableVerticalMerges: CellVerticalMergeType[] = [];
	currentVerticalMerge: CellVerticalMergeType = null;
	tableCellPositions: CellPos[] = [];
	currentCellPosition: CellPos = null;
	// w:tblCellMar of the table being rendered - it is the default margin of its
	// CELLS, so it has to travel down to them (see renderTable)
	tableCellMargins: Record<string, string>[] = [];
	currentCellMargins: Record<string, string> = null;

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
	// the document's saved page breaks were removed (ignoreLastRenderedPageBreak
	// under measured pagination): section starts are then the renderer's to
	// break, see sectionStartsPage
	savedBreaksRemoved = false;
	// measured pagination: the section of each data-docx-sect index, for the
	// page numbers fillPageFields writes
	sectionsByIndex: SectionProperties[] = [];

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
		this.defaultParagraphStyleId = null;
		this.savedBreaksRemoved = false;
		this.numberingLevels = {};
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
			this.defaultParagraphStyleId = document.stylesPart.styles.find(s => s.isDefault && s.target == "p")?.id ?? null;
		}

		// numbering rules go BEFORE the style rules: ECMA-376 17.7.2 applies the
		// numbering a style brings, then the paragraph style, then the direct
		// formatting - numbering in the paragraph's own properties included. Of
		// two equally specific rules the later one wins, so a style beats the
		// numbering it brings itself: a Ril "a)" item is indented by its style
		// (709/284 twips), not by its level (1020/510). Numbering of the
		// paragraph's own gets a more specific rule (see renderNumbering).
		if (document.numberingPart) {
			this.prodessNumberings(document.numberingPart.domNumberings);
			for (const num of document.numberingPart.domNumberings)
				(this.numberingLevels[num.id] ??= {})[num.level] = num;

			result.push(...await this.renderNumbering(document.numberingPart.domNumberings));
			//result.push(...await this.renderNumbering2(document.numberingPart.domNumberings));
		}

		if (document.stylesPart != null)
			result.push(...this.renderStyles(document.stylesPart.styles));

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
		// every render pass numbers the document from its start
		this.numberingCounters = {};

		this.processElement(document);
		const sections = this.splitBySection(document.children, document.props);
		const pages = this.groupByPageBreaks(sections);
		let prevProps = null;
		const sectIndex = new Map<SectionProperties, number>();
		this.sectionsByIndex = [];

		for (let i = 0, l = pages.length; i < l; i++) {
			this.currentFootnoteIds = [];

			const section = pages[i][0];
			let props = section.sectProps;
			this.currentPageIndex = result.length; // global page ordinal (mirror margins)
			const pageElement = this.createPageElement(this.className, props, document.cssStyle);

			// stamp the Word section index (pages of one section share sectProps)
			// so a consumer's page-number stamper can count SECTIONPAGES
			if (this.options.measuredPagination) {
				if (!sectIndex.has(props)) {
					sectIndex.set(props, sectIndex.size);
					this.sectionsByIndex.push(props);
				}
				pageElement.setAttribute("data-docx-sect", String(sectIndex.get(props)));
			}

			this.options.renderHeaders && this.renderHeaderFooter(props.headerRefs, props,
				result.length, prevProps != props, pageElement);

			let contentElement: Node = null;
			for (const sect of pages[i]) {
				// measured pagination paginates a sheet's first article only: once
				// the saved page breaks are gone, a continuous section in the same
				// columns continues in that article, as its text does in Word
				if (contentElement && this.savedBreaksRemoved && this.sameColumns(props, sect.sectProps)) {
					this.renderElements(sect.elements, contentElement);
				} else {
					contentElement = this.createSectionContent(sect.sectProps);
					this.renderElements(sect.elements, contentElement);
					pageElement.appendChild(contentElement);
				}
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
		// sheets as its own breaks dictate, then measure the overflow). A caller
		// that ignores the document's own markers has them removed first, so the
		// pages follow the measurement alone.
		const markersRemoved = this.options.ignoreLastRenderedPageBreak;
		if (markersRemoved)
			removeSavedPageBreaks(body);
		this.options.ignoreLastRenderedPageBreak = false;
		this.savedBreaksRemoved = markersRemoved;

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
			applyMeasuredGeometry(host, this.measuredMiMap, this.hfPartsByPath, settings, lineMul, shrink,
				this.options.explicitSingleLines);
			// tabs move text, so they are laid out before lines are counted
			if (this.options.tabStops)
				resolveTabStops(host, defaultTabPoints(settings));
			// Word's space between paragraphs, before the pages are measured
			if (this.options.addUpParagraphSpacing && settings?.doNotUseHTMLParagraphAutoSpacing)
				addUpParagraphSpacing(host);

			const breaks = findMeasuredBreaks(host, markersRemoved);
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
		this.tableCellMargins = [];
		this.currentCellMargins = null;
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
					// content at the bottom of the reserved band, like Word.
					// A variant this page does not use stays hidden - showing it
					// again would stack two footers and, for a consumer that
					// paginates by height, eat the space of the taller one.
					if (el.style.display !== "none") el.style.display = "flex";
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

	/* Splits a paragraph at the lastRenderedPageBreak marker it carries into
	   [head, tail] clones - the marker itself belongs to neither. Null when it
	   carries none. */
	splitParagraphAtBreak(p: OpenXmlElement): [OpenXmlElement, OpenXmlElement] | null {
		const children = p?.children ?? [];
		let rBreakIndex = -1;
		const pBreakIndex = children.findIndex(r => {
			rBreakIndex = r.children?.findIndex(this.isPageBreakElement.bind(this)) ?? -1;
			return rBreakIndex != -1;
		});
		if (pBreakIndex < 0) return null;
		const breakRun = children[pBreakIndex];
		const head = { ...p, children: children.slice(0, pBreakIndex) } as OpenXmlElement;
		// a continuation must not repeat its number nor count again
		const tail = { ...p, numbering: null, $splitCont: true, children: children.slice(pBreakIndex + 1) } as OpenXmlElement;
		const runChildren = breakRun.children ?? [];
		if (rBreakIndex > 0)
			head.children = head.children.concat([{ ...breakRun, children: runChildren.slice(0, rBreakIndex) }]);
		if (rBreakIndex < runChildren.length - 1)
			tail.children = ([{ ...breakRun, children: runChildren.slice(rBreakIndex + 1) }] as OpenXmlElement[])
				.concat(tail.children);
		return [head, tail];
	}

	// splits a table into row segments at cell-internal page breaks; every
	// segment is a clone keeping the columns (widths). Returns [table] when there
	// is nothing to split.
	splitTableByRowBreaks(table: WmlTable): WmlTable[] {
		const rows = table.children ?? [];
		const breakRows: number[] = [];
		for (let ri = 0; ri < rows.length; ri++)
			if (this.rowHasPageBreak(rows[ri]) || (rows[ri] as any).$splitKeep) breakRows.push(ri);
		if (!breakRows.length) return [table];

		// A row the measured pass SPLIT sits on both pages: the blocks that fit
		// close the page it started on, the rest open the next one. $splitKeep[c]
		// is how many blocks of cell c stayed.
		const partRow = (tr: OpenXmlElement, head: boolean): OpenXmlElement => {
			const keep = (tr as any).$splitKeep as number[];
			const cuts = ((tr as any).$splitCut ?? []) as (number | null)[];
			return {
				...tr,
				children: (tr.children ?? []).map((cell, c) => {
					const n = keep[c] ?? 0;
					const blocks = cell.children ?? [];
					// the block that straddles the boundary carries a break marker
					// at the character the measuring pass cut it at
					const cut = cuts[c] != null ? this.splitParagraphAtBreak(blocks[n]) : null;
					return {
						...cell,
						children: head
							? blocks.slice(0, n).concat(cut ? [cut[0]] : [])
							: (cut ? [cut[1]] : []).concat(blocks.slice(cut ? n + 1 : n))
					};
				})
			} as OpenXmlElement;
		};

		// Word and LibreOffice repeat the w:tblHeader rows on every page the table
		// continues on; the measured pass laid the table out that way, so the
		// segments have to carry them too or the pages come out shorter than
		// measured. The rows are rendered a second time, not copied - nothing in
		// the model addresses a row by identity.
		let headerRows: OpenXmlElement[] = [];
		for (const r of rows) {
			if (!(r as WmlTableRow).isHeader) break;
			headerRows.push(r);
		}

		const bounds = [0, ...breakRows, rows.length];
		const segs: WmlTable[] = [];
		for (let bi = 0; bi < bounds.length - 1; bi++) {
			let seg = rows.slice(bounds[bi], bounds[bi + 1]);
			if (!seg.length) continue;
			// a split row opens this segment with its remainder ...
			if (bi > 0 && (rows[bounds[bi]] as any).$splitKeep)
				seg = [partRow(rows[bounds[bi]], false), ...seg.slice(1)];
			// ... and closes it with the part that still fitted
			const endRow = rows[bounds[bi + 1]];
			if (endRow && (endRow as any).$splitKeep)
				seg = seg.concat([partRow(endRow, true)]);
			if (bi > 0 && headerRows.length && bounds[bi] >= headerRows.length)
				seg = headerRows.concat(seg);
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
				const s = this.findStyle(this.paragraphStyleName(elem as WmlParagraph));

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

		for (let i = 0; i < sections.length; i++) {
			const s = sections[i];
			current.push(s);

			if (this.options.ignoreLastRenderedPageBreak || s.pageBreak || this.isPageBreakSection(prev, s.sectProps)
				|| this.sectionStartsPage(s, sections[i + 1]))
				result.push(current = []);

			prev = s.sectProps;
		}

		return result.filter(x => x.length > 0);
	}

	/* Whether the Word section that follows `s` starts on a new page. Word starts
	   every section on a new page except a continuous one on the same page size
	   (nextPage is the default; evenPage and oddPage may add a blank page, which
	   is not reproduced). A document Word saved has a lastRenderedPageBreak at
	   the top of such a page, which breaks it here; once those markers are
	   removed, the section's type has to. */
	sectionStartsPage(s: Section, next: Section): boolean {
		return this.savedBreaksRemoved && !!next && next.sectProps !== s.sectProps
			&& (next.sectProps?.type !== SectionType.Continuous || this.isPageBreakSection(s.sectProps, next.sectProps));
	}

	// the same column layout - all of a section that createSectionContent uses
	sameColumns(a: SectionProperties, b: SectionProperties): boolean {
		const cols = (p: SectionProperties) => (p?.columns?.numberOfColumns ?? 1) > 1
			? `${p.columns.numberOfColumns}|${p.columns.space}|${!!p.columns.separator}` : "1";
		return cols(a) === cols(b);
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
.${c} { color: black; hyphens: auto; text-underline-position: from-font; text-rendering: geometricPrecision; font-kerning: none; }
section.${c} { box-sizing: border-box; display: flex; flex-flow: column nowrap; position: relative; overflow: hidden; }
section.${c}>article { margin-bottom: auto; z-index: 1; }
/* The space AFTER the last block on a sheet is not carried across the break -
   Word does not let it push the page taller, and neither may we: the article's
   auto margin resolves to 0 as soon as the content plus that space exceeds the
   body band, and the sheet grows instead. Sheets would then differ in size. */
section.${c}>article>*:last-child { margin-bottom: 0; }
section.${c}>header { z-index: 0; }
section.${c}>footer { z-index: 0; }
.${c} table { border-collapse: collapse; }
.${c} table td, .${c} table th { vertical-align: top; }
.${c} td, .${c} th { padding: 0 5.4pt; }
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

		if (this.options.addUpParagraphSpacing) {
			styleText += `.${c} ${PREV_AFTER_RESET}\n`;
		}

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
				var directMinWidth: string = null;
				if (num.suff == null || num.suff == "tab") {
					// the hanging indent the paragraph really gets: its linked
					// style's own beats the level's (the style rules win, see
					// render) - but not in a paragraph numbered by its own w:numPr
					var linkedPara = linkedStyle?.styles?.find(s => s.target == "p")?.values;
					beforeValues["display"] = "inline-block";
					beforeValues["min-width"] = this.markerWidth(linkedPara?.["text-indent"] ?? num.pStyle?.["text-indent"]);
					directMinWidth = this.markerWidth(num.pStyle?.["text-indent"]);
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

				// numberingAsText: the same box around a real text span the
				// paragraph renders itself - no generated content, no CSS counter
				if (this.options.numberingAsText) {
					const labelValues = { ...beforeValues };
					delete labelValues["content"];
					delete labelValues["counter-increment"];
					styleText += this.styleToString(`${selector} > [data-docx-num-label]`, labelValues);
				} else {
					styleText += this.styleToString(`${selector}:before`, beforeValues);
				}
				if (directMinWidth && directMinWidth != beforeValues["min-width"])
					styleText += this.styleToString(`${selector}.${this.numberingDirectClass()}`
						+ (this.options.numberingAsText ? " > [data-docx-num-label]" : ":before"), { "min-width": directMinWidth });
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
			// numbering of the paragraph's own comes with its direct formatting,
			// after its style (ECMA-376 17.7.2): a rule more specific than the style's
			if (num.pStyle && Object.keys(num.pStyle).length)
				styleText += this.styleToString(`${selector}.${this.numberingDirectClass()}`, num.pStyle);
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
			// A w:style without a w:styleId cannot be referenced by anything, and
			// Word ignores it. We must too: with no id it takes the wrapper's own
			// class name (processStyleName(null)), so it would apply itself to
			// every paragraph and every run in the document.
			if (style.id == null && !style.isDocDefaults) {
				if (this.options.debug)
					console.warn(`Skipping style without id: ${style.name}`);
				continue;
			}

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

				// a character style beats the run properties of the default
				// paragraph style, which every paragraph without a w:pStyle carries
				// too (p.<default> span) - in Word the run's own style wins
				if (style.target == "span" && defautStyles["p"])
					selector += `, p.${defautStyles["p"].cssName} ${selector}`;

				if (style.target != subStyle.target)
					selector += ` ${subStyle.target}`;

				if (defautStyles[style.target] == style)
					selector = `.${this.className} ${style.target}, ` + selector;

				styleText += this.styleToString(selector, subStyle.values);

				// The default RUN properties are what a PARAGRAPH MARK is formatted
				// with too, and a paragraph holding nothing but its mark is still a
				// line tall. A browser has no mark, so an empty paragraph would take
				// its size from the user agent instead of from the document. Put the
				// defaults on the wrapper, where anything without a run of its own
				// inherits them. (docDefaults carries no id, and isDefault is not
				// set on it either, hence its own flag.)
				if (style.isDocDefaults && subStyle.target == "span")
					styleText += this.styleToString(`.${this.className}`, subStyle.values);
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

		// without a w:pStyle of its own the paragraph gets the default style's
		// class too: its run properties apply to the paragraph's spans
		const styleName = this.paragraphStyleName(elem);
		if (styleName && !elem.styleName)
			result.classList.add(this.processStyleName(styleName));

		const style = this.findStyle(styleName);
		// tabStops reads the style chain itself (paragraphTabs): elem.tabs stays
		// the paragraph's own stops
		if (!this.options.tabStops)
			elem.tabs ??= style?.paragraphProps?.tabs;  //TODO

		// a continuation of a paragraph split at a measured page break must not
		// repeat its number nor re-increment the numbering counter
		const numbering = (this.options.measuredPagination && (elem as any).$splitCont)
			? null
			: (elem.numbering ?? style?.paragraphProps?.numbering);

		if (numbering) {
			result.classList.add(this.numberingClass(numbering.id, numbering.level));
			// numbered by its own w:numPr, not by its style: the level's paragraph
			// properties beat the style's (see renderNumbering)
			if (numbering === elem.numbering)
				result.classList.add(this.numberingDirectClass());

			// numberingAsText: the label as the paragraph's first span, styled
			// by the marker rule renderNumbering writes for it
			const label = this.options.numberingAsText ? this.numberingLabel(numbering.id, numbering.level) : null;
			if (label != null) {
				const span = this.h({ tagName: "span", children: [label] }) as HTMLElement;
				span.setAttribute("data-docx-num-label", "");
				result.insertBefore(span, result.firstChild);
			}
		}

		if (this.options.measuredPagination && (elem as any).$mi != null)
			result.setAttribute("data-docx-mi", String((elem as any).$mi));

		// w:keepNext - the paginator keeps this paragraph on the page the next
		// block starts on. Resolved here because the style chain lives here: in
		// practice it is a heading style that carries it, not the paragraph.
		if (this.options.measuredPagination && (elem.keepNext ?? style?.paragraphProps?.keepNext))
			result.setAttribute("data-docx-keep-next", "1");

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

	// UAX#14 LB21 ("× HY") forbids a break BEFORE a hyphen, so a run of hyphens
	// is unbreakable and carries the character in front of it - which is what
	// LibreOffice's line breaker (its own ICU tailoring, i18npool line.txt) and
	// Word do. Blink offers a break after every hyphen instead, so a row of
	// dashes - a filler our documents use a lot - gets torn apart and every line
	// after it lands somewhere else than in the printed document. Keep those runs
	// together; the text itself is untouched, so copy, search and the CID
	// anchors are unaffected.
	static readonly HYPHEN_RUN = /.?[-‐]{2,}/g;

	renderText(elem: WmlText) {
		const text = elem.text;
		if (!text || !/[-‐]{2,}/.test(text))
			return this.h(text);

		const parts: any[] = [];
		let last = 0;
		HtmlRenderer.HYPHEN_RUN.lastIndex = 0;
		for (let m: RegExpExecArray; (m = HtmlRenderer.HYPHEN_RUN.exec(text));) {
			if (m.index > last) parts.push(this.h(text.substring(last, m.index)));
			parts.push(this.h({ tagName: "span", style: { "white-space": "nowrap" }, children: [this.h(m[0])] }));
			last = m.index + m[0].length;
		}
		if (last < text.length) parts.push(this.h(text.substring(last)));
		return parts;
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

		if (this.options.tabStops && this.options.measuredPagination) {
			// laid out on its paragraph's stops once measured (resolveTabStops);
			// the em space holds its place until then
			const paragraph = findParent<WmlParagraph>(elem, DomType.Paragraph);
			tabSpan.setAttribute("data-docx-tab", tabStopsAttr(paragraph ? this.paragraphTabs(paragraph) : []));
		} else if (this.options.experimental) {
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
		// the field's first result run shows the whole filled value
		if (elem.fieldResultRest && this.options.measuredPagination)
			return null;

		let children = this.renderElements(elem.children);

		if (elem.verticalAlign) {
			children = [this.h({ tagName: elem.verticalAlign, children: this.renderElements(elem.children) })];
		}

		const result = this.toHTML(elem, ns.html, "span", children);

		if (elem.id)
			result.id = elem.id;

		// mark field result runs so measured pagination can fill the live value
		// (see fillPageFields)
		if (elem.fieldValue && this.options.measuredPagination)
			this.markField(result, elem.fieldValue, elem.fieldArg);

		return result;
	}

	markField(el: HTMLElement, keyword: string, arg: string) {
		el.setAttribute("data-docx-field", keyword);
		if (arg) el.setAttribute("data-docx-field-arg", arg);
	}

	// A simple field (w:fldSimple). Only the fields fillPageFields fills are
	// rendered, and only under measuredPagination (where the live value can be
	// computed) - matching the default renderer, which does not render simple
	// fields at all. The cached result is rendered as a placeholder and then
	// overwritten by fillPageFields.
	renderSimpleField(elem: WmlFieldSimple) {
		if (!elem.fieldValue || !this.options.measuredPagination)
			return null;

		const result = this.toHTML(elem, ns.html, "span", this.renderElements(elem.children));
		this.markField(result, elem.fieldValue, elem.fieldArg);
		return result;
	}

	// Fills PAGE/NUMPAGES/SECTIONPAGES/PAGEREF/REF field placeholders with live
	// values now that the final pages are known. SECTIONPAGES counts pages
	// sharing a section (data-docx-sect, stamped per page in measured mode).
	// PAGEREF shows the page a bookmark starts on, REF the number label of the
	// bookmarked paragraph (numberingAsText) - both the way a table of contents
	// uses them. A field whose bookmark is missing keeps its cached result.
	// PAGE and PAGEREF show the number Word gives a page: numbering runs on
	// across sections unless one restarts it (w:pgNumType w:start), in each
	// section's own format (w:fmt); NUMPAGES counts the pages.
	fillPageFields(pages: HTMLElement[]) {
		const sectKey = (s: HTMLElement, i: number) => s.getAttribute("data-docx-sect") ?? ("solo" + i);
		const perSect: Record<string, number> = {};
		pages.forEach((s, i) => { const k = sectKey(s, i); perSect[k] = (perSect[k] ?? 0) + 1; });
		const pageNumbers: string[] = [];
		let pageNumber = 0;
		pages.forEach((s, i) => {
			const sect = s.getAttribute("data-docx-sect");
			const props = this.sectionsByIndex[parseInt(sect, 10)];
			const start = props?.pageNumber?.start;
			const restarts = (i == 0 || sect != pages[i - 1].getAttribute("data-docx-sect")) && Number.isInteger(start);
			pageNumber = restarts ? start : pageNumber + 1;
			pageNumbers.push(this.formatNumber(pageNumber, props?.pageNumber?.format ?? "decimal"));
		});
		// bookmarks render as <span id>; the first occurrence is where one starts,
		// on the page of the first text after it: a bookmark Word saved in front
		// of a page's first line stays in the empty paragraph head that the saved
		// page break leaves on the page before (splitBySection)
		const textFollows = (b: HTMLElement, page: HTMLElement) => {
			const walker = document.createTreeWalker(page, NodeFilter.SHOW_TEXT);
			walker.currentNode = b;
			for (let n = walker.nextNode(); n; n = walker.nextNode())
				if (n.nodeValue.trim()) return true;
			return false;
		};
		const bookmarks: Record<string, { el: HTMLElement; page: number }> = {};
		pages.forEach((s, i) => s.querySelectorAll<HTMLElement>("span[id]").forEach(b => {
			if (!bookmarks[b.id])
				bookmarks[b.id] = { el: b, page: (textFollows(b, s) || i == pages.length - 1) ? i : i + 1 };
		}));
		const refValue = (keyword: string, arg: string): string | null => {
			const b = bookmarks[arg];
			if (!b) return null;
			if (keyword === "PAGEREF") return pageNumbers[b.page];
			const label = b.el.closest("p")?.querySelector("[data-docx-num-label]");
			return label ? label.textContent.trim().replace(/\.+$/, "") : null;
		};
		pages.forEach((s, i) => {
			const vals: Record<string, string> = {
				PAGE: pageNumbers[i],
				NUMPAGES: String(pages.length),
				SECTIONPAGES: String(perSect[sectKey(s, i)])
			};
			s.querySelectorAll<HTMLElement>("[data-docx-field]").forEach(el => {
				const keyword = el.getAttribute("data-docx-field");
				const v = el.hasAttribute("data-docx-field-arg")
					? refValue(keyword, el.getAttribute("data-docx-field-arg"))
					: vals[keyword];
				if (v != null) el.textContent = v;
			});
		});
	}

	renderTable(elem: WmlTable) {
		this.tableCellPositions.push(this.currentCellPosition);
		this.tableVerticalMerges.push(this.currentVerticalMerge);
		this.tableCellMargins.push(this.currentCellMargins);
		this.currentVerticalMerge = {};
		this.currentCellPosition = { col: 0, row: 0 };
		// w:tblCellMar sets the default margin of the table's CELLS, but on the
		// table itself it is parsed as padding on the table element, where CSS
		// never applies it to them. Take it off the table and hand it to
		// renderTableCell. (From a table STYLE it already arrives as a td rule -
		// the style parser files tblPr under the "td" target.)
		this.currentCellMargins = this.takeCellMargins(elem.cssStyle);

		const children = [];

		if (elem.columns)
			children.push(this.renderTableColumns(elem.columns));

		children.push(...this.renderElements(elem.children));

		this.currentVerticalMerge = this.tableVerticalMerges.pop();
		this.currentCellPosition = this.tableCellPositions.pop();
		this.currentCellMargins = this.tableCellMargins.pop();
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

		const result = this.toHTML(elem, ns.html, "tr", children) as HTMLTableRowElement;

		// w:tblHeader - the measured paginator repeats these at the top of every
		// page the table continues on, the way Word and LibreOffice do
		if (elem.isHeader)
			result.setAttribute("data-docx-thead", "1");

		return result;
	}

	renderTableCellPlaceholder(colSpan: number) {
		return this.h({ tagName: "td", colSpan, style: { border: "none" } });
	}

	/* Moves the padding-* entries out of a table's own properties - that is where
	   its w:tblCellMar lands. Returns null when there are none. */
	takeCellMargins(values: Record<string, string>): Record<string, string> {
		if (!values) return null;
		let out: Record<string, string> = null;
		for (const side of ["top", "right", "bottom", "left"]) {
			const key = "padding-" + side;
			if (values[key] == null) continue;
			(out ??= {})[key] = values[key];
			delete values[key];
		}
		return out;
	}

	renderTableCell(elem: WmlTableCell) {
		// the table's w:tblCellMar is the default for its cells; the cell's own
		// w:tcMar (parsed into its cssStyle) wins
		if (this.currentCellMargins)
			elem.cssStyle = { ...this.currentCellMargins, ...(elem.cssStyle ?? {}) };

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

	/* The style a paragraph is formatted with: its w:pStyle, else the default
	   paragraph style - Word applies that one to every paragraph without a
	   w:pStyle, so the document defaults alone would put its runs in the wrong
	   font. */
	paragraphStyleName(elem: WmlParagraph): string {
		return elem.styleName ?? this.defaultParagraphStyleId;
	}

	/* A paragraph's tab stops: its style's, from the base style down, then its
	   own. A "clear" stop removes the one an earlier level set at that position
	   (a footer style clearing its base style's centre and right stops). */
	paragraphTabs(elem: WmlParagraph): ParagraphTab[] {
		const levels: ParagraphTab[][] = [elem.tabs ?? []];
		const seen: Record<string, boolean> = {};
		for (let style = this.findStyle(this.paragraphStyleName(elem)); style && !seen[style.id]; style = this.findStyle(style.basedOn)) {
			seen[style.id] = true;
			levels.unshift(style.tabs ?? []);
		}
		let stops: ParagraphTab[] = [];
		for (const level of levels) {
			for (const t of level) {
				stops = stops.filter(s => parseFloat(s.position as string) !== parseFloat(t.position as string));
				if (t.style !== "clear") stops.push(t);
			}
		}
		return stops;
	}

	numberingDirectClass() {
		return `${this.className}-num-direct`;
	}

	// the marker box of a "tab" suffix spans the hanging indent; levels
	// declaring hanging="0" parse to "-0.00pt" - only a real (negative)
	// hanging indent replaces the default tab stop
	markerWidth(hangingIndent: string): string {
		return (hangingIndent && parseFloat(hangingIndent) < 0)
			? hangingIndent.substring(1)
			: (this.document.settingsPart?.settings?.defaultTabStop ?? "36pt");
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

	/* The label of the next paragraph at this level, counted the way Word
	   counts: the level goes up by one, every level below it starts over
	   (unless its w:lvlRestart says otherwise), and each %N in the level text
	   shows level N's count in level N's own format. Null when the level has
	   no text (a picture bullet, or no level at all) - those stay CSS. */
	numberingLabel(numId: string, level: number): string {
		const levels = this.numberingLevels[numId];
		const lvl = levels?.[level];
		if (lvl?.levelText == null || lvl.bullet)
			return null;

		const counts = (this.numberingCounters[numId] ??= []);
		counts[level] = (counts[level] ?? lvl.start - 1) + 1;
		for (const key in levels) {
			const below = levels[key];
			// absent w:lvlRestart: restarts after any level above it; 0: never
			if (below.level > level && level < (below.restart ?? below.level))
				counts[below.level] = undefined;
		}

		return lvl.levelText.replace(/%(\d)/g, (s, n) => {
			const ref = levels[parseInt(n, 10) - 1];
			return ref ? this.formatNumber(counts[ref.level] ?? ref.start, ref.format) : "";
		});
	}

	/* A count in a w:numFmt the way Word writes it - letters repeat past z
	   (y, z, aa, bb); formats not covered here fall back to decimal. */
	formatNumber(value: number, format: string): string {
		switch (format) {
			case "none":
			case "bullet":
				return "";

			case "decimalZero":
				return value >= 0 && value < 10 ? "0" + value : String(value);

			case "lowerLetter":
			case "upperLetter": {
				if (value < 1) return String(value);
				const letters = String.fromCharCode(97 + (value - 1) % 26).repeat(Math.floor((value - 1) / 26) + 1);
				return format == "upperLetter" ? letters.toUpperCase() : letters;
			}

			case "lowerRoman":
			case "upperRoman": {
				if (value < 1 || value > 3999) return String(value);
				const digits: [number, string][] = [[1000, "M"], [900, "CM"], [500, "D"], [400, "CD"], [100, "C"],
					[90, "XC"], [50, "L"], [40, "XL"], [10, "X"], [9, "IX"], [5, "V"], [4, "IV"], [1, "I"]];
				let roman = "";
				for (const [n, s] of digits)
					while (value >= n) { roman += s; value -= n; }
				return format == "lowerRoman" ? roman.toLowerCase() : roman;
			}

			default:
				return String(value);
		}
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

