/* Word's paragraph spacing, laid out on the rendered pages (option addUpParagraphSpacing).

   A document that turns HTML paragraph spacing off (w:doNotUseHTMLParagraphAutoSpacing,
   which Word writes into its documents) has Word ADD a paragraph's space after and
   the next paragraph's space before. CSS collapses the two margins to the larger
   one - measured in Word's own PDFs: 6 pt after and 12 pt before make 18 pt in such
   a document and 12 pt in one without the setting.

   addUpParagraphSpacing gives each paragraph that follows another one in the flow
   that one's space after on top of its own space before, through
   --docx-prev-after; the collapse then yields the sum. A paragraph that starts its
   container - a cell, or a sheet the measured pagination has just opened - has no
   space after above it on that page, and PREV_AFTER_RESET (a rule of the default
   style) takes the variable back to 0 there.

   Runs in measured pagination before the page breaks are found, and again on the
   attached pages. */

export const PREV_AFTER_RESET = "p:first-child { --docx-prev-after: 0px !important; }";

export function addUpParagraphSpacing(root: HTMLElement) {
	// read every margin first, then write: one style computation for the whole pass
	const changes: { p: HTMLElement; before: number; prevAfter: number }[] = [];
	const paragraphs = Array.from(root.querySelectorAll<HTMLElement>("p"));
	for (let i = 0; i < paragraphs.length; i++) {
		const p = paragraphs[i];
		const prev = previousInFlow(p);
		if (!prev || prev.tagName !== "P") continue;
		const prevAfter = parseFloat(getComputedStyle(prev).marginBottom) || 0;
		const before = parseFloat(getComputedStyle(p).marginTop) || 0;
		// with either one 0 the larger one already is the sum
		if (prevAfter > 0 && before > 0) changes.push({ p, before, prevAfter });
	}
	for (let i = 0; i < changes.length; i++) {
		const c = changes[i];
		c.p.style.setProperty("--docx-prev-after", `${c.prevAfter}px`);
		c.p.style.marginTop = `calc(${c.before}px + var(--docx-prev-after))`;
	}
}

// The sibling whose bottom margin meets this element's top margin: hidden, floating
// and absolutely positioned siblings (framed paragraphs) are out of the flow.
function previousInFlow(el: Element): Element | null {
	for (let prev = el.previousElementSibling; prev; prev = prev.previousElementSibling) {
		const cs = getComputedStyle(prev);
		if (cs.display !== "none" && cs.float === "none" && cs.position !== "absolute" && cs.position !== "fixed")
			return prev;
	}
	return null;
}
