import { ParagraphProperties, ParagraphTab } from "./paragraph";
import { RunProperties } from "./run";

export interface IDomStyle {
    id: string;
    name?: string;
    cssName?: string;
    aliases?: string[];
    target: string;
    basedOn?: string;
    isDefault?: boolean;
    isDocDefaults?: boolean;
    styles: IDomSubStyle[];
    linked?: string;
    next?: string;

    paragraphProps: ParagraphProperties;
    runProps: RunProperties;
    // the style's own tab stops (its w:tabs, "clear" entries included). The
    // renderer merges paragraphProps with the base style's entry by entry,
    // overwriting them, so the chain is resolved from these instead.
    tabs?: ParagraphTab[];
}

export interface IDomSubStyle {
    target: string;
	mod?: string;
    values: Record<string, string>;
}