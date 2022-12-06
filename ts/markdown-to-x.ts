import { marked } from "marked";
import { escape, unescape } from "lodash";

const separator = "\t";

const block = (blokType: string, f: (...a: any) => string) => (text: string) =>
	`\n${blokType}${separator}In\n` +
	f(text) +
	`\n${blokType}${separator}Out\n\n`;
const line = (lineType: string) => (text: string) =>
	`${lineType}${separator}${text}\n`;
const tableLine = (lineType: string) => (header: string, body: string) =>
	`${lineType}${separator}${header}${body}\n`;
const inline = (inlineType: string) => (text: string) =>
	`<${inlineType}>${text}</${inlineType}>`;
const newline = () => "\n";
const empty = () => "";

const DocxRenderer: marked.Renderer = {
	// Block elements
	code: block("code", (s: string) => escape(s)),
	blockquote: block("blockquote", (s: string) => s),
	html: empty,
	heading: (text, section) =>
		"section" + separator + section.toString() + separator + text + "\n\n", //block,
	hr: newline,
	list: block("List", (f) => f),
	listitem: line("listitem"),
	checkbox: empty,
	paragraph: block("paragraph", (f) => f),
	table: tableLine("table"),
	tablerow: block("tablerow", (f) => f),
	tablecell: block("tablecell", (f) => f),

	// Inline elements
	strong: inline("strong"),
	em: inline("em"),
	codespan: inline("codespan"),
	br: newline,
	del: inline("del"),
	link: (_0, _1, text) => `${_0}:${_1}:${text}`,
	image: (_0, _1, text) => "image" + separator + _0,
	text: (text) => text,
	// etc.
	options: {},
};

/**
 * Converts markdown to plaintext using the marked Markdown library.
 * Accepts [MarkedOptions](https://marked.js.org/using_advanced#options) as
 * the second argument.
 *
 * NOTE: The output of markdownToTxt is NOT sanitized. The output may contain
 * valid HTML, JavaScript, etc. Be sure to sanitize if the output is intended
 * for web use.
 *
 * @param markdown the markdown text to txtify
 * @param options  the marked options
 * @returns the unmarked text
 */
export function markdownToDocx(
	markdown: string,
	options?: marked.MarkedOptions
): string {
	const unmarked = marked(markdown, { ...options, renderer: DocxRenderer });
	const unescaped = unescape(unmarked);
	const trimmed = unescaped.trim();
	return trimmed;
}

export default markdownToDocx;
