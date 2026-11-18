import { DocumentParser } from "../document-parser";
import { Length } from "../document/common";
import { XmlParser } from "../parser/xml-parser";

export interface WmlSettings {
	defaultTabStop: Length;
	footnoteProps: NoteProperties;
	endnoteProps: NoteProperties;
	autoHyphenation: boolean;
	evenAndOddHeaders: boolean;
	mirrorMargins: boolean;
	compatibilityMode: number;
	// Word ADDS a paragraph's space after and the next one's space before, instead of
	// taking the larger of the two (the HTML rule)
	doNotUseHTMLParagraphAutoSpacing: boolean;
}

export interface NoteProperties {
	nummeringFormat: string;
	defaultNoteIds: string[];
}

export function parseSettings(elem: Element, xml: XmlParser) {
	var result = {} as WmlSettings; 

	for (let el of xml.elements(elem)) {
		switch(el.localName) {
			case "defaultTabStop": result.defaultTabStop = xml.lengthAttr(el, "val"); break;
			case "footnotePr": result.footnoteProps = parseNoteProperties(el, xml); break;
			case "endnotePr": result.endnoteProps = parseNoteProperties(el, xml); break;
			case "autoHyphenation": result.autoHyphenation = xml.boolAttr(el, "val", true); break;
			case "evenAndOddHeaders": result.evenAndOddHeaders = xml.boolAttr(el, "val", true); break;
			case "mirrorMargins": result.mirrorMargins = xml.boolAttr(el, "val", true); break;
			case "compat":
				result.compatibilityMode = parseCompatibilityMode(el, xml);
				result.doNotUseHTMLParagraphAutoSpacing = parseCompatFlag(el, xml, "doNotUseHTMLParagraphAutoSpacing");
				break;
		}
	}

    return result;
}

export function parseCompatFlag(elem: Element, xml: XmlParser, localName: string): boolean {
	const flag = xml.element(elem, localName);
	return !!flag && xml.boolAttr(flag, "val", true);
}

export function parseCompatibilityMode(elem: Element, xml: XmlParser): number {
	for (let el of xml.elements(elem, "compatSetting")) {
		if (xml.attr(el, "name") == "compatibilityMode")
			return xml.intAttr(el, "val");
	}

	return null;
}

export function parseNoteProperties(elem: Element, xml: XmlParser) {
	var result = {
		defaultNoteIds: []
	} as NoteProperties; 

	for (let el of xml.elements(elem)) {
		switch(el.localName) {
			case "numFmt": 
				result.nummeringFormat = xml.attr(el, "val");
				break;

			case "footnote": 
			case "endnote": 
				result.defaultNoteIds.push(xml.attr(el, "id"));
				break;
		}
	}

    return result;
}