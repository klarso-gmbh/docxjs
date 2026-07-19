/*
 * Measured pagination for docx-preview (Options.measuredPagination).
 *
 * docx-preview can only break pages at markers it finds in the document model
 * (explicit page breaks, section ends, Word's <w:lastRenderedPageBreak/>).
 * Documents exported without those markers render as one endless sheet. This
 * module reproduces Word's page breaks by MEASURING a first render pass, then
 * inserting lastRenderedPageBreak markers into the model at the measured break
 * points so a second render pass splits the pages natively - which means the
 * library's own per-page header/footer selection and mirrored-margin geometry
 * apply automatically.
 *
 * It also applies the layout metrics Word uses but CSS does not express:
 *  - "auto" line spacing measured from the font's natural line height,
 *  - justified inter-word space shrinking (Word 2013+),
 *  - automatic hyphenation via a caller-provided callback.
 *
 * The heavy DOM logic here (break finding by measurement) reproduces the page
 * layout established empirically against Word-made PDFs; the constants and
 * rules are calibrated, not derived.
 */

import { DomType, OpenXmlElement, WmlText, WmlBreak } from './document/dom';
import { WmlRun } from './document/run';
import { WmlParagraph } from './document/paragraph';
import { DocumentElement } from './document/document';
import { WmlSettings } from './settings/settings';
import { WordDocument } from './word-document';

const EPS = 0.5;
const SAFETY_MAX_SHEETS = 10000;
const MIN_LINES_KEEP = 2; // orphan control: lines that must stay on the sheet
const MIN_LINES_MOVE = 2; // widow control: lines that must move together
// Word 2013+ justified line breaking shrinks inter-word spaces to pull a word
// up. Calibrated against Word-made PDFs (sweeping -25%..-45% for whole-document
// line-break parity); -40% of the space advance also absorbs Word's slightly
// narrower (unkerned, GDI-rounded) line measure. Overridable via
// Options.justificationShrink (fraction, e.g. 0.4).
const JUSTIFY_SHRINK_DEFAULT = -0.40;
const SHY = String.fromCharCode(0xAD); // soft hyphen
// letters incl. Latin-1/Latin-Extended; Word only hyphenates longer words
const HYPH_WORD_RE = /[A-Za-zÀ-ɏ]{5,}/g;

function px(v: string): number {
	const n = parseFloat(v);
	return isNaN(n) ? 0 : n;
}

// ---------------------------------------------------------------------------
// Break descriptors recorded while measuring, then applied to the model.
// ---------------------------------------------------------------------------

type BreakDesc =
	| { kind: "before"; mi: number }                     // page break before a block
	| { kind: "tableRow"; mi: number; row: number }      // split a table before an (absolute) row
	| { kind: "paraChar"; mi: number; offset: number };  // split a paragraph at a character offset

interface PaginateCtx {
	breaks: BreakDesc[];
	record(desc: BreakDesc): void;
}

// ===========================================================================
// 1. Model indexing - so measured DOM blocks map back to model elements.
// ===========================================================================

/* Assigns a document-order index ($mi) to every body-level paragraph and
   table (recursing into table cells so cell-internal breaks resolve too), and
   returns the mi -> element map. The renderer stamps the index as
   data-docx-mi. */
export function assignModelIndices(body: DocumentElement): Map<number, OpenXmlElement> {
	const map = new Map<number, OpenXmlElement>();
	let next = 0;
	const walk = (el: OpenXmlElement) => {
		if (!el.children) return;
		for (const c of el.children) {
			if (c.type == DomType.Paragraph || c.type == DomType.Table) {
				(c as any).$mi = next;
				map.set(next, c);
				next++;
			}
			if (c.type == DomType.Table || c.type == DomType.Row || c.type == DomType.Cell)
				walk(c);
		}
	};
	walk(body);
	return map;
}

// ===========================================================================
// 2. Geometry - Word line metrics + justified shrink, measured once per font.
// ===========================================================================

/* The document-default "auto" line multiplier (styles.xml docDefaults,
   w:line/240). Returns null when absent or when the default rule is
   exact/atLeast (docx-preview already emits those as fixed line-heights). */
export function getAutoLineMul(document: WordDocument): number | null {
	const styles = document.stylesPart?.styles;
	if (!styles) return null;
	// docDefaults are parsed as a style with id == null holding a "p" target
	const def = styles.find(s => s.id == null && s.styles?.some(x => x.target == "p"));
	const p = def?.styles?.find(x => x.target == "p");
	const lh = p?.values?.["line-height"];
	if (!lh) return null;
	// auto => a plain unitless number; exact/atLeast => "..pt" or "calc(..)"
	if (/[a-z(%]/i.test(lh)) return null;
	const n = parseFloat(lh);
	return isNaN(n) ? null : n;
}

/* Maps every rendered header/footer <p> back to its source part paragraph, so
   measured geometry written to the part model is re-emitted on the second
   render pass (header/footer parts are rendered from their own model, not the
   indexed body). Correlates by document order within each part instance. */
function buildHeaderFooterMap(host: HTMLElement, hfPartsByPath: Record<string, OpenXmlElement>): Map<HTMLElement, OpenXmlElement> {
	const map = new Map<HTMLElement, OpenXmlElement>();
	const partParas: Record<string, OpenXmlElement[]> = {};
	const parasOf = (el: OpenXmlElement): OpenXmlElement[] => {
		const out: OpenXmlElement[] = [];
		const walk = (e: OpenXmlElement) => {
			if (!e.children) return;
			for (const c of e.children) {
				if (c.type == DomType.Paragraph) out.push(c);
				walk(c);
			}
		};
		walk(el);
		return out;
	};
	const instances = host.querySelectorAll<HTMLElement>("[data-docx-part]");
	instances.forEach(inst => {
		const path = inst.getAttribute("data-docx-part");
		const root = hfPartsByPath[path];
		if (!root) return;
		const paras = (partParas[path] ??= parasOf(root));
		const domPs = inst.querySelectorAll<HTMLElement>("p");
		for (let i = 0; i < domPs.length && i < paras.length; i++)
			map.set(domPs[i], paras[i]);
	});
	return map;
}

/* Applies Word-faithful line metrics and justified shrink to the model (so
   both render passes agree) and to the measuring host (so this pass measures
   with them). Also stamps per-paragraph widow/orphan + keep-lines data
   attributes the break finder reads. */
export function applyMeasuredGeometry(
	host: HTMLElement,
	miMap: Map<number, OpenXmlElement>,
	hfPartsByPath: Record<string, OpenXmlElement>,
	settings: WmlSettings,
	lineMul: number | null,
	shrink: number = JUSTIFY_SHRINK_DEFAULT
) {
	const justifyShrink = !!(settings && settings.compatibilityMode >= 15);
	if (!lineMul && !justifyShrink) return;

	const cache: Record<string, number> = {};
	// Probe the font's natural line as a ratio of the em size by INHERITANCE
	// inside a real document span - canvas and inline-style font probes get
	// silently substituted in some engines. Probe large and scale down so
	// integer ascent/descent rounding is negligible; skip hosts that are not
	// laid out (hidden header/footer variants measure 0x0).
	function probeFace(faceKey: string, hostEl: HTMLElement): number | null {
		if (cache[faceKey] !== undefined) return cache[faceKey];
		const probe = document.createElement("span");
		probe.style.cssText = "position:absolute;visibility:hidden;display:block;font-size:256px;line-height:normal";
		probe.textContent = "Hg";
		hostEl.appendChild(probe);
		const h = probe.getBoundingClientRect().height;
		probe.remove();
		if (h < 128) return null;
		cache[faceKey] = h / 256;
		return cache[faceKey];
	}
	function probeWidth(hostEl: HTMLElement, text: string): number {
		const probe = document.createElement("span");
		probe.style.cssText = "position:absolute;visibility:hidden;white-space:pre;font-size:256px";
		probe.textContent = text;
		hostEl.appendChild(probe);
		const w = probe.getBoundingClientRect().width;
		probe.remove();
		return w;
	}
	function probeSpace(spaceKey: string, hostEl: HTMLElement): number | null {
		if (cache[spaceKey] !== undefined) return cache[spaceKey];
		const w = probeWidth(hostEl, "x x") - probeWidth(hostEl, "xx");
		if (w <= 0) return null;
		cache[spaceKey] = w / 256;
		return cache[spaceKey];
	}

	const hfMap = buildHeaderFooterMap(host, hfPartsByPath);

	// pass A: resolve every paragraph's face and probe where possible
	interface Item { p: HTMLElement; model: OpenXmlElement; size: number; family: string; hasOwnLine: boolean; spaceKey: string | null; }
	const items: Item[] = [];
	const faceOfFamily: Record<string, Record<string, boolean>> = {};
	const paras = host.querySelectorAll<HTMLElement>("section.docx p");
	for (let i = 0; i < paras.length; i++) {
		const p = paras[i];
		const miAttr = p.getAttribute("data-docx-mi");
		let model: OpenXmlElement = null;
		if (miAttr != null) model = miMap.get(parseInt(miAttr, 10));
		else model = hfMap.get(p);
		if (!model) continue; // unmapped (e.g. notes) - leave to defaults

		// per-paragraph widow/orphan + keep-lines for the splitter (body only)
		if (miAttr != null) {
			const mp = model as WmlParagraph;
			if (mp.widowControl === false) p.setAttribute("data-docx-widow-control", "0");
			if (mp.keepLines === true) p.setAttribute("data-docx-keep-lines", "1");
		}

		// an inline line-height is a document rule (exact/atLeast) - keep it,
		// but justified spaces still shrink in Word
		const hasOwnLine = !!p.style.lineHeight;
		const span = p.querySelector("span");
		const cs = getComputedStyle(span || p);
		const faceKey = cs.fontFamily + "|" + cs.fontWeight + "|" + cs.fontStyle;
		if (!hasOwnLine) {
			probeFace(faceKey, span || p);
			(faceOfFamily[cs.fontFamily] = faceOfFamily[cs.fontFamily] || {})[faceKey] = true;
		}
		const shrink = justifyShrink && getComputedStyle(p).textAlign === "justify";
		if (shrink) probeSpace("sp:" + faceKey, span || p);
		if (!hasOwnLine || shrink)
			items.push({ p, model, size: parseFloat(cs.fontSize), family: cs.fontFamily, hasOwnLine, spaceKey: shrink ? "sp:" + faceKey : null });
	}

	// pass B: one line ratio per FAMILY = the minimum of its probed faces.
	// Office fonts share line metrics across faces; engines only inflate
	// synthesized faces, so the minimum is the un-synthesized truth.
	const familyRatio: Record<string, number> = {};
	for (const fam in faceOfFamily) {
		let min: number = null;
		for (const key in faceOfFamily[fam])
			if (cache[key] !== undefined && (min === null || cache[key] < min)) min = cache[key];
		familyRatio[fam] = min !== null ? min : 1.2;
	}

	for (const it of items) {
		const css = (it.model.cssStyle = it.model.cssStyle || {});
		if (!it.hasOwnLine) {
			if (lineMul) {
				const lh = (lineMul * it.size * familyRatio[it.family]).toFixed(2) + "px";
				css["line-height"] = lh;
				it.p.style.lineHeight = lh;
			}
			// align the paragraph strut with its runs so the page default font
			// cannot inflate any line box
			css["font-size"] = it.size + "px";
			css["font-family"] = it.family;
			it.p.style.fontSize = it.size + "px";
			it.p.style.fontFamily = it.family;
		}
		if (it.spaceKey && cache[it.spaceKey] !== undefined) {
			const ws = (shrink * cache[it.spaceKey] * it.size).toFixed(2) + "px";
			css["word-spacing"] = ws;
			it.p.style.wordSpacing = ws;
		}
	}
}

// ===========================================================================
// 3. Break finding by measurement (DOM port of the calibrated paginator).
// ===========================================================================

function sheetGeometry(section: HTMLElement) {
	const cs = getComputedStyle(section);
	let target = px(cs.minHeight);
	if (target <= 0) target = px(cs.height);
	return { target, padBottom: px(cs.paddingBottom) };
}

function contentRoot(section: HTMLElement): HTMLElement {
	for (let i = 0; i < section.children.length; i++)
		if (section.children[i].tagName === "ARTICLE") return section.children[i] as HTMLElement;
	return section;
}

/* Children after the <article> (footers) eat into usable height only by their
   real flow contribution (the reserved band has an equal negative margin, so a
   footer that fits contributes zero). */
function postArticleHeight(section: HTMLElement): number {
	let h = 0, seenArticle = false;
	for (let i = 0; i < section.children.length; i++) {
		const c = section.children[i] as HTMLElement;
		if (c.tagName === "ARTICLE") { seenArticle = true; continue; }
		if (!seenArticle) continue;
		const flow = c.offsetHeight + (parseFloat(getComputedStyle(c).marginBottom) || 0);
		if (flow > 0) h += flow;
	}
	return h;
}

function limitY(section: HTMLElement, geo: { target: number; padBottom: number }): number {
	return section.getBoundingClientRect().top + geo.target - geo.padBottom - postArticleHeight(section);
}

/* Empty continuation sheet: same shell, cloned headers/footers reset to the
   default variant (so it measures with real capacity, not the tall first-page
   header), empty article, inserted after `section`. */
function newSheet(section: HTMLElement): HTMLElement {
	const clone = section.cloneNode(false) as HTMLElement;
	clone.classList.add("docx-cont-page");
	let art: HTMLElement = null;
	for (let i = 0; i < section.children.length; i++) {
		const c = section.children[i];
		if (c.tagName === "ARTICLE") { art = c.cloneNode(false) as HTMLElement; clone.appendChild(art); }
		else clone.appendChild(c.cloneNode(true));
	}
	if (!art) { art = document.createElement("article"); clone.appendChild(art); }
	for (let j = 0; j < clone.children.length; j++) {
		const hf = clone.children[j] as HTMLElement;
		const type = hf.getAttribute && hf.getAttribute("data-hf-type");
		if (!type) continue;
		hf.style.display = (type !== "default") ? "none" : (hf.tagName === "FOOTER" ? "flex" : "");
	}
	section.parentNode.insertBefore(clone, section.nextSibling);
	return clone;
}

function moveFollowing(toParent: HTMLElement, el: Element) {
	while (el) { const next = el.nextElementSibling; toParent.appendChild(el); el = next; }
}

function splitTableAt(table: HTMLTableElement, splitIdx: number): HTMLTableElement {
	const cont = table.cloneNode(false) as HTMLTableElement;
	for (let i = 0; i < table.children.length; i++) {
		if (table.children[i].tagName === "COLGROUP") { cont.appendChild(table.children[i].cloneNode(true)); break; }
	}
	const body = document.createElement("tbody");
	cont.appendChild(body);
	while (table.rows.length > splitIdx) body.appendChild(table.rows[splitIdx]); // .rows is live
	const base = parseInt(table.getAttribute("data-docx-rowbase") || "0", 10) || 0;
	cont.setAttribute("data-docx-rowbase", String(base + splitIdx));
	return cont;
}

function imageAtomOf(img: Element, root: Element): Element {
	let a = img;
	while (a.parentElement && a.parentElement !== root
		&& a.parentElement.textContent.trim() === ""
		&& a.parentElement.querySelectorAll("img").length === 1)
		a = a.parentElement;
	return a;
}

function atomsOf(el: HTMLElement): (Text | Element)[] {
	const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT, {
		acceptNode(n: Node) {
			if (n.nodeType === 3) return (n as Text).nodeValue.length ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_REJECT;
			return (n as Element).tagName === "IMG" ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_SKIP;
		}
	});
	const atoms: (Text | Element)[] = [];
	let n: Node;
	while ((n = walker.nextNode())) {
		if (n.nodeType === 3) { atoms.push(n as Text); continue; }
		const atom = imageAtomOf(n as Element, el);
		if (atoms.indexOf(atom as any) < 0) atoms.push(atom as Element);
	}
	return atoms;
}

function isImageAtom(a: Text | Element): a is Element {
	return a.nodeType === 1;
}

interface LineBox { top: number; bottom: number; imageOnly: boolean; }

function lineBoxesOf(atoms: (Text | Element)[]): LineBox[] {
	const rects: { top: number; bottom: number; image: boolean }[] = [];
	const range = document.createRange();
	for (const atom of atoms) {
		if (isImageAtom(atom)) {
			const ir = atom.getBoundingClientRect();
			if (ir.height > 0 && ir.width > 0) rects.push({ top: ir.top, bottom: ir.bottom, image: true });
			continue;
		}
		range.selectNodeContents(atom);
		const rl = range.getClientRects();
		for (let j = 0; j < rl.length; j++)
			if (rl[j].height > 0 && rl[j].width > 0) rects.push({ top: rl[j].top, bottom: rl[j].bottom, image: false });
	}
	rects.sort((a, b) => a.top - b.top);
	const lines: LineBox[] = [];
	for (const rc of rects) {
		const last = lines[lines.length - 1];
		if (last && rc.top - last.top <= 1.5) {
			if (rc.bottom > last.bottom) last.bottom = rc.bottom;
			if (!rc.image) last.imageOnly = false;
		} else {
			lines.push({ top: rc.top, bottom: rc.bottom, imageOnly: rc.image });
		}
	}
	return lines;
}

function charRect(node: Text, i: number): DOMRect {
	const r = document.createRange();
	r.setStart(node, i);
	r.setEnd(node, i + 1);
	return r.getBoundingClientRect();
}

interface LinePos { node?: Text; offset?: number; before?: Element; atomIndex: number; }

function lineStartPosition(atoms: (Text | Element)[], lineTop: number): LinePos | null {
	for (let i = 0; i < atoms.length; i++) {
		const atom = atoms[i];
		if (isImageAtom(atom)) {
			if (atom.getBoundingClientRect().top >= lineTop - 1.5) return { before: atom, atomIndex: i };
			continue;
		}
		const len = atom.nodeValue.length;
		if (charRect(atom, len - 1).top < lineTop - 1.5) continue;
		let lo = 0, hi = len - 1, found = -1;
		while (lo <= hi) {
			const mid = (lo + hi) >> 1;
			if (charRect(atom, mid).top >= lineTop - 1.5) { found = mid; hi = mid - 1; }
			else lo = mid + 1;
		}
		if (found >= 0) return { node: atom, offset: found, atomIndex: i };
	}
	return null;
}

/* Splits a multi-line paragraph at the last line boundary that fits, honoring
   widow/orphan minimums. Records the model break point (paragraph mi + absolute
   character offset) and mirrors the DOM split so subsequent measurement of the
   continuation is correct. Returns the continuation sheet or null. */
function trySplitParagraph(section: HTMLElement, p: HTMLElement, limit: number, mustSplit: boolean, ctx: PaginateCtx): HTMLElement | null {
	if (p.tagName !== "P") return null;
	if (!mustSplit && p.getAttribute("data-docx-keep-lines") === "1") return null;
	try {
		const atoms = atomsOf(p);
		if (!atoms.length) return null;
		const lines = lineBoxesOf(atoms);
		if (lines.length < 2) return null;
		const relax = mustSplit || p.getAttribute("data-docx-widow-control") === "0";
		const minKeep = relax ? 1 : MIN_LINES_KEEP;
		const minMove = relax ? 1 : MIN_LINES_MOVE;
		let fit = -1;
		for (let i = 0; i < lines.length; i++) { if (lines[i].bottom <= limit + EPS) fit = i; else break; }
		if (lines.length < minKeep + minMove) return null;
		if (fit + 1 < minKeep) return null;
		if (lines.length - (fit + 1) < minMove) { fit = lines.length - minMove - 1; if (fit + 1 < minKeep) return null; }
		if (fit >= lines.length - 1) return null;
		const pos = lineStartPosition(atoms, lines[fit + 1].top);
		if (!pos || (pos.atomIndex === 0 && (pos.before || pos.offset === 0))) return null;

		// absolute break offset within the paragraph: text counts by character
		// (INCLUDING soft hyphens - hyphenation is in the model text too, so the
		// measured DOM offset maps directly), each image/atomic box counts as 1
		// slot (so a break between two stacked images maps to the model too).
		let abs = pos.before ? 0 : pos.offset;
		for (let n = 0; n < pos.atomIndex; n++)
			abs += isImageAtom(atoms[n]) ? 1 : (atoms[n] as Text).nodeValue.length;

		const range = document.createRange();
		range.selectNodeContents(p);
		if (pos.before) range.setStartBefore(pos.before);
		else range.setStart(pos.node, pos.offset);

		const cont = p.cloneNode(false) as HTMLElement; // same classes/styles
		// numbering renders via ::before + CSS counters; a continuation must not
		// repeat the number nor count again
		cont.className = Array.prototype.filter.call(cont.classList, (c: string) => c.indexOf("-num-") < 0).join(" ");
		cont.appendChild(range.extractContents());
		const base = parseInt(p.getAttribute("data-docx-split-offset") || "0", 10) || 0;
		cont.setAttribute("data-docx-split-offset", String(base + abs));

		const mi = miOf(p);
		if (mi != null) ctx.record({ kind: "paraChar", mi, offset: base + abs });

		const sheet = newSheet(section);
		const contArt = contentRoot(sheet);
		contArt.appendChild(cont);
		moveFollowing(contArt, p.nextElementSibling);
		return sheet;
	} catch (e) {
		return null;
	}
}

function miOf(el: Element): number | null {
	const v = el && el.getAttribute && el.getAttribute("data-docx-mi");
	return v == null ? null : parseInt(v, 10);
}

/* Break after an element that can't be split and doesn't fit anywhere. */
function breakAfter(section: HTMLElement, el: Element, ctx: PaginateCtx): HTMLElement | null {
	const next = el.nextElementSibling;
	if (!next) return null;
	const nmi = miOf(next);
	if (nmi != null) ctx.record({ kind: "before", mi: nmi });
	const sheet = newSheet(section);
	moveFollowing(contentRoot(sheet), next);
	return sheet;
}

function paginateSheet(section: HTMLElement, geo: { target: number; padBottom: number }, ctx: PaginateCtx): HTMLElement | null {
	const art = contentRoot(section);
	const limit = limitY(section, geo);

	let crossing: HTMLElement = null;
	for (let i = 0; i < art.children.length; i++) {
		if ((art.children[i] as HTMLElement).getBoundingClientRect().bottom > limit + EPS) { crossing = art.children[i] as HTMLElement; break; }
	}
	if (!crossing) return null;
	const isFirst = (crossing === art.firstElementChild);

	if (crossing.tagName === "TABLE") {
		const table = crossing as HTMLTableElement;
		const rows = table.rows;
		const rowbase = parseInt(table.getAttribute("data-docx-rowbase") || "0", 10) || 0;
		const tmi = miOf(table);
		let splitIdx = -1;
		for (let j = 0; j < rows.length; j++) {
			if (rows[j].getBoundingClientRect().bottom > limit + EPS) { splitIdx = j; break; }
		}
		if (splitIdx < 0) splitIdx = rows.length;
		if (splitIdx === 0) {
			if (!isFirst) {
				// starts too low: move the whole table (and the rest) over
				if (tmi != null) ctx.record(rowbase > 0 ? { kind: "tableRow", mi: tmi, row: rowbase } : { kind: "before", mi: tmi });
				const sheetA = newSheet(section);
				moveFollowing(contentRoot(sheetA), crossing);
				return sheetA;
			}
			if (rows.length <= 1) return breakAfter(section, crossing, ctx); // single monster row
			splitIdx = 1;
		}
		if (splitIdx >= rows.length && rows.length > 0 && !crossing.nextElementSibling) return null;
		if (tmi != null && splitIdx < rows.length) ctx.record({ kind: "tableRow", mi: tmi, row: rowbase + splitIdx });
		const sheetB = newSheet(section);
		const contArt = contentRoot(sheetB);
		if (splitIdx < rows.length) contArt.appendChild(splitTableAt(table, splitIdx));
		moveFollowing(contArt, crossing.nextElementSibling);
		return sheetB;
	}

	const split = trySplitParagraph(section, crossing, limit, isFirst, ctx);
	if (split) return split;
	if (isFirst) return breakAfter(section, crossing, ctx); // unsplittable block taller than a page
	const cmi = miOf(crossing);
	if (cmi != null) ctx.record({ kind: "before", mi: cmi });
	const sheet = newSheet(section);
	moveFollowing(contentRoot(sheet), crossing);
	return sheet;
}

function needsPagination(section: HTMLElement, geo: { target: number }): boolean {
	const tolerance = Math.max(24, geo.target * 0.05);
	return section.getBoundingClientRect().height > geo.target + tolerance;
}

/* Runs the measured paginator over the host and returns the model break
   descriptors (the DOM is mutated in place - that host is discarded). */
export function findMeasuredBreaks(host: HTMLElement): BreakDesc[] {
	const ctx: PaginateCtx = { breaks: [], record(d) { this.breaks.push(d); } };
	const originals = Array.prototype.slice.call(host.querySelectorAll("section.docx")) as HTMLElement[];
	let safety = SAFETY_MAX_SHEETS;
	for (const orig of originals) {
		const geo = sheetGeometry(orig);
		if (!geo.target || !needsPagination(orig, geo)) continue;
		let sheet: HTMLElement = orig;
		while (sheet && safety-- > 0) sheet = paginateSheet(sheet, geo, ctx);
	}
	return ctx.breaks;
}

// ===========================================================================
// 4. Applying the breaks to the model as lastRenderedPageBreak markers.
// ===========================================================================

function makeBreakRun(): WmlRun {
	return { type: DomType.Run, children: [{ type: DomType.Break, break: "lastRenderedPageBreak" } as WmlBreak] } as unknown as WmlRun;
}

// A break position in the model text: text atoms count by character (soft
// hyphens included), each image/tab/symbol counts as one slot - matching the
// offsets the measured DOM produced (see trySplitParagraph).
interface Atom { kind: "text" | "slot"; topChild: OpenXmlElement; leaf: OpenXmlElement; start: number; len: number; }

/* Flattens a paragraph's break atoms in document order, recording for each its
   direct paragraph child (the level splitBySection acts on) and start offset.
   Skips runs the renderer drops (field code, deleted text). */
function paragraphAtoms(p: WmlParagraph): { atoms: Atom[]; total: number } {
	const atoms: Atom[] = [];
	let pos = 0;
	for (const topChild of p.children ?? []) {
		const visit = (el: OpenXmlElement) => {
			if (el.type == DomType.Run && ((el as WmlRun).fieldRun || (el as WmlRun).fieldCodeContent)) return;
			if (el.type == DomType.DeletedText) return; // renderChanges off in this path
			if (el.type == DomType.Text) {
				const len = (el as WmlText).text?.length ?? 0;
				atoms.push({ kind: "text", topChild, leaf: el, start: pos, len });
				pos += len;
				return;
			}
			if (el.type == DomType.Image || el.type == DomType.Tab || el.type == DomType.Symbol) {
				atoms.push({ kind: "slot", topChild, leaf: el, start: pos, len: 1 });
				pos += 1;
				return;
			}
			for (const c of el.children ?? []) visit(c);
		};
		visit(topChild);
	}
	return { atoms, total: pos };
}

// the ancestor of `leaf` that is a direct child of `parent`
function directChildOf(parent: OpenXmlElement, leaf: OpenXmlElement): OpenXmlElement {
	let c = leaf;
	while (c.parent && c.parent !== parent) c = c.parent;
	return c;
}

/* Inserts a lastRenderedPageBreak into paragraph `p` at break offset `off`,
   splitting the containing run like Word (between two stacked images, or at a
   character within a text run) so splitBySection breaks the page there. */
function insertParagraphBreak(p: WmlParagraph, off: number) {
	const { atoms, total } = paragraphAtoms(p);
	if (!atoms.length) { p.children = [makeBreakRun(), ...(p.children ?? [])]; return; }
	if (off >= total) {
		const lastTop = atoms[atoms.length - 1].topChild;
		const i = p.children.indexOf(lastTop);
		p.children.splice(i < 0 ? p.children.length : i + 1, 0, makeBreakRun());
		return;
	}
	if (off < 0) off = 0;
	const A = atoms.find(a => off < a.start + a.len) ?? atoms[atoms.length - 1];
	const pi = p.children.indexOf(A.topChild);
	if (pi < 0) { p.children.unshift(makeBreakRun()); return; }
	const local = off - A.start;
	const directChild = directChildOf(A.topChild, A.leaf);

	// character break inside a text run: run(..left) | breakRun | run2(right..)
	if (A.kind === "text" && local > 0 && local < A.len
		&& A.topChild.type == DomType.Run && directChild === A.leaf) {
		const run = A.topChild as WmlRun;
		const text = A.leaf as WmlText;
		const ci = run.children.indexOf(text);
		const left = { ...text, text: text.text.substring(0, local) } as WmlText;
		const right = { ...text, text: text.text.substring(local) } as WmlText;
		const run2 = { ...run, children: [right as OpenXmlElement, ...run.children.slice(ci + 1)] } as WmlRun;
		(run as any).children = [...run.children.slice(0, ci), left];
		p.children.splice(pi + 1, 0, makeBreakRun(), run2);
		return;
	}

	// boundary before `directChild` within a run (e.g. between two drawings)
	if (A.topChild.type == DomType.Run) {
		const run = A.topChild as WmlRun;
		const ci = run.children.indexOf(directChild);
		if (ci <= 0) {
			p.children.splice(pi, 0, makeBreakRun()); // break before the whole run
		} else {
			const run2 = { ...run, children: run.children.slice(ci) } as WmlRun;
			(run as any).children = run.children.slice(0, ci);
			p.children.splice(pi + 1, 0, makeBreakRun(), run2);
		}
		return;
	}

	// topChild is not a run (hyperlink/smartTag): break before it (approximate)
	p.children.splice(pi, 0, makeBreakRun());
}

/* Applies recorded break descriptors to the model. Paragraph splits at the same
   paragraph are applied high-offset-first so earlier offsets don't shift. */
export function applyBreaksToModel(breaks: BreakDesc[], miMap: Map<number, OpenXmlElement>) {
	// group paragraph char breaks by mi, dedup others
	const paraOffsets: Record<number, number[]> = {};
	const seen = new Set<string>();
	for (const b of breaks) {
		if (b.kind === "paraChar") { (paraOffsets[b.mi] ??= []).push(b.offset); continue; }
		const key = b.kind + ":" + b.mi + ":" + (b.kind === "tableRow" ? b.row : "");
		if (seen.has(key)) continue;
		seen.add(key);
		if (b.kind === "before") {
			const el = miMap.get(b.mi);
			if (el) (el as any).$pageBreakBefore = true;
		} else if (b.kind === "tableRow") {
			markTableRowBreak(miMap.get(b.mi), b.row);
		}
	}
	for (const miStr in paraOffsets) {
		const el = miMap.get(parseInt(miStr, 10)) as WmlParagraph;
		if (!el) continue;
		const offs = Array.from(new Set(paraOffsets[miStr])).sort((a, b) => b - a); // high first
		for (const off of offs) insertParagraphBreak(el, off);
	}
}

/* Marks a table so splitBySection starts a new page at (absolute) row `row`:
   inserts a lastRenderedPageBreak into the first paragraph of that row's first
   cell (where Word places it) - splitBySection detects it and splits the table. */
function markTableRowBreak(table: OpenXmlElement, row: number) {
	if (!table || !table.children || row <= 0 || row >= table.children.length) return;
	const tr = table.children[row];
	const cell = tr?.children?.[0];
	if (!cell) return;
	let firstPara = cell.children?.find(c => c.type == DomType.Paragraph) as WmlParagraph;
	if (!firstPara) {
		firstPara = { type: DomType.Paragraph, children: [] } as any;
		cell.children.unshift(firstPara);
	}
	firstPara.children = [makeBreakRun(), ...(firstPara.children ?? [])];
}

// ===========================================================================
// 5. Hyphenation (model text), before the first render pass.
// ===========================================================================

/* Collects unique body words (letters, length >= 5), asks the callback to
   hyphenate them, and inserts the returned soft hyphens into the model text so
   both render passes are consistent. The document bytes are never touched. */
export async function hyphenateModel(body: DocumentElement, hyphenate: (words: string[]) => string[] | Promise<string[]>): Promise<number> {
	const texts: WmlText[] = [];
	const collect = (el: OpenXmlElement) => {
		if (el.type == DomType.Text) { if ((el as WmlText).text) texts.push(el as WmlText); return; }
		if (el.type == DomType.DeletedText) return;
		for (const c of el.children ?? []) collect(c);
	};
	collect(body);

	const seen = new Set<string>();
	const words: string[] = [];
	for (const t of texts) {
		const m = t.text.match(HYPH_WORD_RE);
		if (!m) continue;
		for (const w of m) if (!seen.has(w)) { seen.add(w); words.push(w); }
	}
	if (!words.length) return 0;

	let result: string[];
	try {
		result = await hyphenate(words);
	} catch (e) {
		return 0;
	}
	if (!result) return 0;
	const map: Record<string, string> = Object.create(null);
	for (let i = 0; i < words.length && i < result.length; i++)
		if (result[i] && result[i].indexOf(SHY) >= 0) map[words[i]] = result[i];

	let changed = 0;
	for (const t of texts) {
		const nv = t.text.replace(HYPH_WORD_RE, w => map[w] || w);
		if (nv !== t.text) { t.text = nv; changed++; }
	}
	return changed;
}

// ===========================================================================
// 6. Anchored (framePr) frame positioning - post-attach, per-page.
// ===========================================================================

/* With mirrored margins even pages already had their padding swapped in
   createPageElement. Inside/outside-anchored frames (marked by the renderer via
   --docx-frame-x) go to the OUTSIDE margin of their page, inner edge hSpace from
   the text column: outer inset = margin - frame width - hSpace. */
export function positionAnchoredFrames(container: HTMLElement, settings: WmlSettings) {
	const mirror = !!(settings && settings.mirrorMargins);
	const sheets = container.querySelectorAll<HTMLElement>("section.docx");
	for (let i = 0; i < sheets.length; i++) {
		const even = mirror && (i + 1) % 2 === 0;
		const sectCs = getComputedStyle(sheets[i]);
		const padL = parseFloat(sectCs.paddingLeft) || 0;
		const padR = parseFloat(sectCs.paddingRight) || 0;
		const ps = sheets[i].querySelectorAll<HTMLElement>("p");
		for (let j = 0; j < ps.length; j++) {
			const cs = getComputedStyle(ps[j]);
			if (cs.position !== "absolute") continue;
			const anchor = cs.getPropertyValue("--docx-frame-x").trim();
			if (anchor !== "outside" && anchor !== "inside") continue;
			const toLeft = (anchor === "outside") === even; // outside = right on odd pages
			const hSpacePt = parseFloat(cs.getPropertyValue("--docx-frame-hspace")) || 0;
			const inset = Math.max(0, (toLeft ? padL : padR) - (parseFloat(cs.width) || 0) - hSpacePt * (96 / 72));
			ps[j].style.left = toLeft ? (inset + "px") : "auto";
			ps[j].style.right = toLeft ? "auto" : (inset + "px");
		}
	}
}
