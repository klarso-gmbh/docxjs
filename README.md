[![npm version](https://badge.fury.io/js/docx-preview.svg)](https://www.npmjs.com/package/docx-preview)
[![Support Ukraine](https://img.shields.io/badge/Support-Ukraine-blue?style=flat&logo=adguard)](https://war.ukraine.ua/)

# docxjs
Docx rendering library

Demo - https://volodymyrbaydalka.github.io/docxjs/

Goal
----
Goal of this project is to render/convert DOCX document into HTML document with keeping HTML semantic as much as possible. 
That means library is limited by HTML capabilities (for example Google Docs renders *.docx document on canvas as an image).

Installation
-----
```
npm install docx-preview
```

Usage
-----
```html
<!--lib uses jszip-->
<script src="https://unpkg.com/jszip/dist/jszip.min.js"></script>
<script src="docx-preview.min.js"></script>
<script>
    var docData = <document Blob>;

    docx.renderAsync(docData, document.getElementById("container"))
        .then(x => console.log("docx: finished"));
</script>
<body>
    ...
    <div id="container"></div>
    ...
</body>
```
API
---
```ts
// renders document into specified element
renderAsync(
    document: Blob | ArrayBuffer | Uint8Array, // could be any type that supported by JSZip.loadAsync
    bodyContainer: HTMLElement, //element to render document content,
    styleContainer: HTMLElement, //element to render document styles, numbeings, fonts. If null, bodyContainer will be used.
    options: {
        className: string = "docx", //class name/prefix for default and document style classes
        inWrapper: boolean = true, //enables rendering of wrapper around document content
        hideWrapperOnPrint: boolean = false, //disable wrapper styles on print
        ignoreWidth: boolean = false, //disables rendering width of page
        ignoreHeight: boolean = false, //disables rendering height of page
        ignoreFonts: boolean = false, //disables fonts rendering
        breakPages: boolean = true, //enables page breaking on page breaks
        ignoreLastRenderedPageBreak: boolean = true, //disables page breaking on lastRenderedPageBreak elements
        measuredPagination: boolean = false, //render, measure where content overflows each page, insert lastRenderedPageBreak markers at those points and re-render — breaks pages without markers in the file (see "Measured pagination"). Also applies Word "auto" line metrics, justified-space shrink and hyphenation.
        hyphenate: (words: string[]) => (string[] | Promise<string[]>) = null, //hyphenation callback used by measuredPagination: given the unique words (letters, length >= 5), return each word with soft hyphens (U+00AD) inserted, in the same order
        onMeasuredBreaks: (breaks: { paraIndex: number, offset: number }[]) => void = null, //reports the breaks measuredPagination decided, addressed against the ORIGINAL document: paraIndex = paragraph index in document order, offset = soft-hyphen-free character offset within it (0 for whole-block breaks)
        justificationShrink: number = 0.4, //measuredPagination: fraction by which justified inter-word spaces may shrink (Word 2013+ line-breaking behavior)
        experimental: boolean = false, //enables experimental features (tab stops calculation)
        trimXmlDeclaration: boolean = true, //if true, xml declaration will be removed from xml documents before parsing
        useBase64URL: boolean = false, //if true, images, fonts, etc. will be converted to base 64 URL, otherwise URL.createObjectURL is used
        renderChanges: false, //enables experimental rendering of document changes (inserions/deletions)
        renderHeaders: true, //enables headers rendering
        renderFooters: true, //enables footers rendering
        renderFootnotes: true, //enables footnotes rendering
        renderEndnotes: true, //enables endnotes rendering
        renderComments: false, //enables experimental comments rendering
        renderAltChunks: true, //enables altChunks (html parts) rendering
        debug: boolean = false, //enables additional logging
        h: ({ ns, tagName, className, style, children, ...props } | Node | string): Node, //experimental hook for HTML rendering, default implementation - defaultOptions.h
    }): Promise<WordDocument>

defaultOptions: Options; // default options

/// ==== experimental / internal API ===
// this API could be used to modify document before rendering
// renderAsync = parseAsync + renderDocument

// parse document and return internal document object
parseAsync(
    document: Blob | ArrayBuffer | Uint8Array,
    options: Options
): Promise<WordDocument>

// render internal document object and return list of nodes
renderDocument(
    wordDocument: WordDocument,
    options: Options
): Promise<Node[]>
```

Thumbnails, TOC and etc.
------
Thumbnails is added only for example and it's not part of library. Library renders DOCX into HTML, so it can't be efficiently used for thumbnails. 

Table of contents is built using the TOC fields and there is no efficient way to get table of contents at this point, since fields is not supported yet (http://officeopenxml.com/WPtableOfContents.php)

Breaks
------
The library breaks pages at markers it finds in the document:

- **Manual page break** — `<w:br w:type="page"/>`, inserted by the user.
- **Rendered page break** — `<w:lastRenderedPageBreak/>`, a hint that editors such as MS Word write to record where a page broke the last time the document was laid out. Honored only when `ignoreLastRenderedPageBreak` is `false`; it defaults to `true`, so set it to `false` to break on these.
- **Section / page-setup change** — e.g. a switch from portrait to landscape.

A document exported without any of these markers renders as one continuous sheet. Splitting it into pages means laying the document out and measuring where content overflows each page — not done on every edit, because re-measuring is expensive. The `measuredPagination` option performs that measurement as an opt-in pass.

Measured pagination
------
`measuredPagination: true` breaks pages for documents that carry no break markers. It renders the document once, measures where content overflows each page, inserts `<w:lastRenderedPageBreak/>` markers into the model at those points, and re-renders — so pages split through the normal render path (per-page header/footer selection and mirrored margins included). It also applies the layout metrics CSS does not express by default, so line breaks track a word processor's closely:

- "auto" line spacing measured from the font's natural line height (not `line-height: normal`),
- justified inter-word space shrinking (Word 2013+; tunable via `justificationShrink`),
- automatic hyphenation via the caller-supplied `hyphenate` callback (the document bytes are never modified — soft hyphens live only in the render model),
- live `PAGE` / `NUMPAGES` / `SECTIONPAGES` field values, stamped per page after pagination.

The extra render pass is why it is opt-in. `onMeasuredBreaks` reports the break positions it chose, addressed against the original document, for tooling that wants to persist them back into the file.

Status and stability
------
So far I can't come up with final approach of parsing documents and final structure of API. Only **renderAsync** function is stable and definition shouldn't be changed in future. Inner implementation of parsing and rendering may be changed at any point of time.

Contributing
------
Please do not include contents of `./dist` folder in your PR's. Otherwise I most likely will reject it due to stability and security concerns.
