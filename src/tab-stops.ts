import { ParagraphTab } from "./document/paragraph";
import { WmlSettings } from "./settings/settings";

/* Word tab stops, laid out by measurement (option tabStops).

   The renderer writes each tab as a span marked data-docx-tab, carrying the
   explicit stops of its paragraph. resolveTabStops gives every such span the
   width that takes the text after it to its stop - a left stop starts that
   text there, a right stop ends it there, a centred one centres it (a decimal
   stop is taken as a right one) - and fills it with the stop's leader.

   Word measures stops from the edge of the text column (the page margin, or
   the cell's), not from the paragraph's indent. Past the last explicit stop
   the document's default stops apply, and a hanging indent is a stop of its
   own on the first line. A stop past the paragraph's right edge counts as
   that edge.

   Runs in measured pagination before the page breaks are found, and again on
   the attached pages, where filled fields - the page numbers of a table of
   contents - may have changed the text a right stop aligns. */

const PX_PER_PT = 96 / 72;
const LEADERS: Record<string, string> = { dot: ".", middleDot: "·", hyphen: "-", underscore: "_", heavy: "_" };
const DEFAULT_TAB_PT = 36; // Word's default when the settings name none

interface Stop { pos: number; align: string; leader: string } // pos in px from the column edge

// the stops of a paragraph as the tab span carries them: [position pt, alignment, leader]
export function tabStopsAttr(tabs: ParagraphTab[] | undefined): string {
	const stops = (tabs ?? [])
		.filter(t => t.style !== "clear" && t.style !== "bar") // bar stops draw a line, text skips them
		.map(t => [parseFloat(t.position as string), t.style, t.leader || "none"]);
	return JSON.stringify(stops);
}

export function defaultTabPoints(settings: WmlSettings | undefined): number {
	const pt = parseFloat(settings?.defaultTabStop as string);
	return pt > 0 ? pt : DEFAULT_TAB_PT;
}

function stopsOf(span: HTMLElement): Stop[] {
	let raw: [number, string, string][] = [];
	try { raw = JSON.parse(span.getAttribute("data-docx-tab") || "[]"); } catch { /* malformed: default stops only */ }
	return raw.map(([pt, style, leader]) => ({
		pos: pt * PX_PER_PT,
		align: style === "right" || style === "end" || style === "decimal" ? "right" : style === "center" ? "center" : "left",
		leader
	})).sort((a, b) => a.pos - b.pos);
}

// the stop a tab at x (px from the column edge) moves to
function nextStop(stops: Stop[], x: number, hangingAt: number | null, defaultPx: number): Stop {
	const EPS = 0.5;
	if (hangingAt !== null && x < hangingAt - EPS && !stops.some(s => s.pos > x + EPS && s.pos < hangingAt))
		return { pos: hangingAt, align: "left", leader: "none" };
	const explicit = stops.find(s => s.pos > x + EPS);
	if (explicit) return explicit;
	const last = stops.length ? stops[stops.length - 1].pos : 0;
	const from = Math.max(x, last);
	return { pos: (Math.floor(from / defaultPx + EPS / defaultPx) + 1) * defaultPx, align: "left", leader: "none" };
}

// natural width of the text that follows a tab, up to the next tab of the
// paragraph - summed per text fragment, so text that has wrapped to the next
// line while the tab was still narrow counts in full
function followingWidth(span: HTMLElement, next: HTMLElement | null, p: HTMLElement): number {
	const range = document.createRange();
	range.setStartAfter(span);
	if (next) range.setEndBefore(next); else range.setEnd(p, p.childNodes.length);
	let width = 0;
	const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
	for (let n = walker.nextNode(); n; n = walker.nextNode()) {
		if (!range.intersectsNode(n)) continue;
		const text = document.createRange();
		text.selectNodeContents(n);
		for (const r of Array.from(text.getClientRects())) width += r.width;
	}
	return width;
}

const leaderWidths: Record<string, number> = {};
function charWidth(span: HTMLElement, ch: string): number {
	const cs = getComputedStyle(span);
	const key = ch + "|" + cs.font + "|" + cs.letterSpacing;
	if (leaderWidths[key] === undefined) {
		const probe = document.createElement("span");
		probe.style.cssText = "position:absolute;visibility:hidden;white-space:pre";
		probe.textContent = ch.repeat(20);
		span.appendChild(probe);
		leaderWidths[key] = probe.getBoundingClientRect().width / 20;
		probe.remove();
	}
	return leaderWidths[key];
}

export function resolveTabStops(root: HTMLElement, defaultTabPt: number) {
	const defaultPx = defaultTabPt * PX_PER_PT;
	const spans = Array.from(root.querySelectorAll<HTMLElement>("span[data-docx-tab]"));
	for (let i = 0; i < spans.length; i++) {
		const span = spans[i];
		const p = span.closest("p");
		if (!p) continue;
		const cs = getComputedStyle(p);
		const pRect = p.getBoundingClientRect();
		const margin = parseFloat(cs.marginLeft) || 0;
		const column = pRect.left - margin;
		const indent = margin + (parseFloat(cs.borderLeftWidth) || 0) + (parseFloat(cs.paddingLeft) || 0);
		const hangingAt = parseFloat(cs.textIndent) < 0 ? indent : null;

		// a stop past the paragraph's right edge is taken at that edge: text
		// aligned there would not fit on the line and wrap (Word can let it run
		// into the margin instead, depending on the document's compatibility)
		const right = pRect.right - (parseFloat(cs.borderRightWidth) || 0) - (parseFloat(cs.paddingRight) || 0) - column;
		const x = span.getBoundingClientRect().left - column;
		const stop = nextStop(stopsOf(span), x, hangingAt, defaultPx);
		if (stop.pos > right) stop.pos = Math.max(x, right);
		const next = spans[i + 1] && spans[i + 1].closest("p") === p ? spans[i + 1] : null;
		let width = stop.pos - x;
		if (stop.align !== "left") {
			const after = followingWidth(span, next, p);
			width -= stop.align === "right" ? after : after / 2;
		}
		width = Math.max(0, width);

		// the tab's text is a space - whoever reads the text sees one - and its
		// leader; padding in front makes up the rest of the width. An inline
		// span, so the line keeps the height and baseline its text gives it.
		const leader = LEADERS[stop.leader];
		let text = " ";
		if (leader) {
			const n = Math.floor((width - charWidth(span, " ")) / charWidth(span, leader));
			if (n > 0) text += leader.repeat(n);
		}
		span.style.whiteSpace = "pre";
		span.style.paddingLeft = "0px";
		span.textContent = width >= charWidth(span, " ") ? text : "";
		const own = span.getBoundingClientRect().width;
		span.style.paddingLeft = Math.max(0, width - own).toFixed(2) + "px";
	}
}
