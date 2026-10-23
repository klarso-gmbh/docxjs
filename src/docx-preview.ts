import { WordDocument } from './word-document';
import { DocumentParser } from './document-parser';
import { HtmlRenderer } from './html-renderer';
import { positionAnchoredFrames, pinSheetHeights } from './measured-pagination';
import { h } from './html';

export interface Options {
    inWrapper: boolean;
    hideWrapperOnPrint: boolean;
    ignoreWidth: boolean;
    ignoreHeight: boolean;
    ignoreFonts: boolean;
    breakPages: boolean;
    debug: boolean;
    experimental: boolean;
    className: string;
    trimXmlDeclaration: boolean;
    renderHeaders: boolean;
    renderFooters: boolean;
    renderFootnotes: boolean;
	renderEndnotes: boolean;
    ignoreLastRenderedPageBreak: boolean;
	useBase64URL: boolean;
	renderChanges: boolean;
    renderComments: boolean;
    renderAltChunks: boolean;
    renderHeaderFooterVariants: boolean;
    // When true, the renderer measures the laid-out document, inserts
    // lastRenderedPageBreak markers into the model where the content overflows
    // a page, and re-renders so pages split at Word-faithful boundaries. Also
    // applies Word's "auto" line metrics, justified space shrinking and (with
    // `hyphenate`) automatic hyphenation. Off by default: the render is then
    // byte-identical to stock docx-preview. The lastRenderedPageBreak markers
    // the document was saved with break pages as well only with
    // `ignoreLastRenderedPageBreak: false`; otherwise they are removed first
    // and the measurement alone decides, as Word re-paginates on open.
    measuredPagination: boolean;
    // Batch hyphenation callback used only under measuredPagination when the
    // document opts into automatic hyphenation. Receives the unique body words
    // (length >= 5) and must return each word with U+00AD soft hyphens inserted
    // at its permissible break points (or the word unchanged), in the same
    // order. May return synchronously or as a Promise. docx-preview inserts the
    // soft hyphens into the model before measuring so both render passes agree.
    hyphenate: (words: string[]) => string[] | Promise<string[]>;
    // Called during measured pagination with the break positions the library
    // decided, addressed against the ORIGINAL document: each entry is the
    // document-order index of the target paragraph over all body paragraphs
    // (depth-first, including table-cell paragraphs - the order
    // getElementsByTagNameNS(w,"p") yields) plus a soft-hyphen-free character
    // offset within it (0 for whole-block breaks). Invoked BEFORE the model is
    // mutated to insert the markers, so a caller can write the same
    // lastRenderedPageBreak markers into a pristine copy of word/document.xml
    // (e.g. to compare the measured pagination against a real renderer). Only
    // called under measuredPagination.
    onMeasuredBreaks?: (breaks: { paraIndex: number; offset: number }[]) => void;
    // Under measuredPagination, the fraction by which justified inter-word
    // spaces may shrink to pull a word onto a line (Word 2013+ behavior).
    // Given as a positive fraction (e.g. 0.4 = up to 40%); defaults to 0.4.
    justificationShrink?: number;
    // Under measuredPagination, give each paragraph that no line-spacing rule
    // reaches (none in docDefaults, its style or itself - Word lays it out
    // single) an explicit single line: the em ratio of its font's line
    // metrics, probed large, times the font size. Otherwise the engine's
    // "normal" applies, built from metrics rounded to whole pixels - 3% short
    // per line for some fonts. Meant for pages that load their fonts with
    // Word's line metrics (FontFace ascent/descent/lineGap overrides).
    // Off by default.
    explicitSingleLines?: boolean;
    h: typeof h;
}

export const defaultOptions: Options = {
    ignoreHeight: false,
    ignoreWidth: false,
    ignoreFonts: false,
    breakPages: true,
    debug: false,
    experimental: false,
    className: "docx",
    inWrapper: true,
    hideWrapperOnPrint: false,
    trimXmlDeclaration: true,
    ignoreLastRenderedPageBreak: true,
    renderHeaders: true,
    renderFooters: true,
    renderFootnotes: true,
	renderEndnotes: true,
	useBase64URL: false,
	renderChanges: false,
    renderComments: false,
    renderAltChunks: true,
    renderHeaderFooterVariants: false,
    measuredPagination: false,
    hyphenate: null,
    onMeasuredBreaks: null,
    justificationShrink: 0.4,
    explicitSingleLines: false,
    h: h
};

export function parseAsync(data: Blob | any, userOptions?: Partial<Options>): Promise<any>  {
    const ops = { ...defaultOptions, ...userOptions };
    return WordDocument.load(data, new DocumentParser(ops), ops);
}

export async function renderDocument(document: any, userOptions?: Partial<Options>): Promise<any> {
    const ops = { ...defaultOptions, ...userOptions };
    const renderer = new HtmlRenderer();
    return await renderer.render(document, ops);
}

export async function renderAsync(data: Blob | any, bodyContainer: HTMLElement, styleContainer?: HTMLElement, userOptions?: Partial<Options>): Promise<any> {
	const doc = await parseAsync(data, userOptions);
	const nodes = await renderDocument(doc, userOptions);

    styleContainer ??= bodyContainer;
    styleContainer.innerHTML = "";
    bodyContainer.innerHTML = "";

    for (let n of nodes) {
        const c = n.nodeName === "STYLE" ? styleContainer : bodyContainer;
        c.appendChild(n);
    }

    // Anchored (framePr) drawings that hang into the inside/outside margin need
    // the page's final, possibly mirror-swapped padding to be positioned - which
    // is only measurable once the pages are attached. (The even/odd padding swap
    // itself happens during render, in createPageElement.)
    if (userOptions?.measuredPagination) {
        positionAnchoredFrames(bodyContainer, doc.settingsPart?.settings);
        // last: the anchors are measured against the sheet as laid out
        pinSheetHeights(bodyContainer);
    }

    return doc;
}