import { XmlParser } from "../parser/xml-parser";
import { CommonProperties, parseCommonProperty } from "./common";
import { OpenXmlElement } from "./dom";

export interface WmlRun extends OpenXmlElement, RunProperties {
    id?: string;
    verticalAlign?: string;
	fieldRun?: boolean;
	fieldCodeContent?: boolean;
	// set on the (first) result run of a PAGE/NUMPAGES/SECTIONPAGES/PAGEREF/REF
	// field so the renderer marks it (data-docx-field) and measured pagination
	// fills the live value per page
	fieldValue?: string;
	// the bookmark a PAGEREF/REF field points to
	fieldArg?: string;
	// a further result run of such a field: the filled value replaces it
	fieldResultRest?: boolean;
}

export interface RunProperties extends CommonProperties {

}

export function parseRunProperties(elem: Element, xml: XmlParser): RunProperties {
    let result = <RunProperties>{};

    for(let el of xml.elements(elem)) {
        parseRunProperty(el, result, xml);
    }

    return result;
}

export function parseRunProperty(elem: Element, props: RunProperties, xml: XmlParser) {
    if (parseCommonProperty(elem, props, xml))
        return true;

    return false;
}