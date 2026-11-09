/// <reference types="node" />

import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import ts from "typescript6";
import { docTabToMarkdown } from "../src/markdown-converter";
import {
  DOCS_TYPES_MODULE_PREFIX, DRIVE_TYPES_MODULE_PREFIX, GMAIL_TYPES_MODULE_PREFIX,
  SHEETS_TYPES_MODULE_PREFIX, SLIDES_TYPES_MODULE_PREFIX, stripTypeModulePrefix,
} from "../src/type-bundle";
import { buildTab } from "./doc-fixture";

const SOURCE_DIR = join(dirname(fileURLToPath(import.meta.url)), "../src");

function sourcePath(name: string): string {
  return join(SOURCE_DIR, name);
}

function source(name: string): string {
  return readFileSync(sourcePath(name), "utf8");
}

function compileAgentTypes(sourceText: string): string[] {
  const fileName = "/agent-types.ts";
  const options: ts.CompilerOptions = {
    module: ts.ModuleKind.ESNext,
    moduleResolution: ts.ModuleResolutionKind.Bundler,
    noEmit: true,
    skipLibCheck: true,
    strict: true,
    target: ts.ScriptTarget.ESNext,
  };
  const baseHost = ts.createCompilerHost(options);
  const host: ts.CompilerHost = {
    ...baseHost,
    fileExists: name => name === fileName || baseHost.fileExists(name),
    getSourceFile: (name, languageVersion, onError, shouldCreateNewSourceFile) =>
      name === fileName
        ? ts.createSourceFile(name, sourceText, languageVersion, true)
        : baseHost.getSourceFile(name, languageVersion, onError, shouldCreateNewSourceFile),
    readFile: name => name === fileName ? sourceText : baseHost.readFile(name),
  };
  const program = ts.createProgram([fileName], options, host);
  return ts.getPreEmitDiagnostics(program).map(diagnostic =>
    ts.flattenDiagnosticMessageText(diagnostic.messageText, "\n"));
}

/** Guidance the agent must still be reading when it picks a tab to act on. */
const TAB_GUIDANCE = [
  "Call `listTabs()` before `getContent()`",
  "reads exactly one tab and never combines tabs",
  "pass an ID returned by `listTabs()`",
  "only when `listTabs()` returns exactly one tab",
];

/**
 * Fails to compile unless `GoogleDocTab` has exactly the flattened adjacency-list members.
 *
 * Mutual assignability alone misses an added or removed *optional* property, since excess
 * properties are permitted in both directions, so the key sets are compared as well.
 */
const TAB_SHAPE_CHECK = `
type ExpectedGoogleDocTab = {
  id: string; title: string; parentTabId?: string; index: number; nestingLevel: number;
};
type Mutual<A, B> = [A] extends [B] ? ([B] extends [A] ? true : false) : false;
const tabShapeIsExact: Mutual<GoogleDocTab, ExpectedGoogleDocTab> = true;
const tabKeysAreExact: Mutual<keyof GoogleDocTab, keyof ExpectedGoogleDocTab> = true;
`;

/** The names a declaration declares at top level, which share one scope in a flat bundle. */
function topLevelNames(sourceText: string): string[] {
  const file = ts.createSourceFile("/names.ts", sourceText, ts.ScriptTarget.ESNext);
  return file.statements.flatMap(statement =>
    (ts.isTypeAliasDeclaration(statement) || ts.isInterfaceDeclaration(statement) ||
      ts.isClassDeclaration(statement) || ts.isFunctionDeclaration(statement) ||
      ts.isEnumDeclaration(statement)) && statement.name
      ? [statement.name.text]
      : ts.isVariableStatement(statement)
        ? statement.declarationList.declarations.flatMap(declaration =>
          ts.isIdentifier(declaration.name) ? [declaration.name.text] : [])
        : []);
}

function duplicates(names: string[]): string[] {
  return names.filter((name, i) => names.indexOf(name) !== i);
}

/** The names of the table types a directly bound Doc's write session declares. */
const DOC_TABLE_TYPE_NAMES = ["NewDocTable", "DocTableInsertion", "DocTableDeletion"];

/** The table methods of a directly bound Doc, each with its exact signature. */
const DOC_TABLE_METHODS = [
  ["insertTable", "insertTable(table: NewDocTable, tabId?: string): Promise<number>;"],
  ["insertTableRows", "insertTableRows(rows: DocTableInsertion, tabId?: string): Promise<void>;"],
  [
    "insertTableColumns",
    "insertTableColumns(columns: DocTableInsertion, tabId?: string): Promise<void>;",
  ],
  ["deleteTableRows", "deleteTableRows(rows: DocTableDeletion, tabId?: string): Promise<void>;"],
  [
    "deleteTableColumns",
    "deleteTableColumns(columns: DocTableDeletion, tabId?: string): Promise<void>;",
  ],
  ["deleteTable", "deleteTable(table: number, tabId?: string): Promise<void>;"],
] as const;

/** Every table method called as an agent calls it, and the calls the types must refuse. */
const DOC_TABLE_USAGE = `
declare const doc: GoogleDocSession;
const a: Promise<number> = doc.insertTable(
  { rows: 2, columns: 2, cells: [["Name", "Status"], ["Ana"]], after: "## Results" }, "t.0");
const b: Promise<number> = doc.insertTable({ rows: 3, columns: 4 });
async function big(rows: string[][]) {
  const table = await doc.insertTable(
    { rows: 100, columns: 5, cells: rows.slice(0, 100), after: "## Data" });
  await doc.insertTableRows({ table, cells: rows.slice(100) });
}
const c: Promise<void> = doc.insertTableRows({ table: 0, cells: [["Fix login", "Ana", "Open"]] });
const d: Promise<void> = doc.insertTableRows({ table: 0, at: 1, count: 2 }, "t.0");
const e: Promise<void> = doc.insertTableColumns({ table: 1, at: 0, cells: [["Owner"], ["Ana"]] });
const f: Promise<void> = doc.deleteTableRows({ table: 0, at: 2 });
const g: Promise<void> = doc.deleteTableColumns({ table: 0, at: 2, count: 2 }, "t.0");
const h: Promise<void> = doc.deleteTable(0, "t.0");
const reader: GoogleDocReadSession = doc;
// @ts-expect-error A new table states its size.
doc.insertTable({ cells: [["a"]] });
// @ts-expect-error A new table's position is a number.
const wrong: Promise<void> = doc.insertTable({ rows: 1, columns: 1 });
// @ts-expect-error A table is named by its position.
doc.insertTableRows({ at: 0 });
// @ts-expect-error A deletion names where it starts.
doc.deleteTableRows({ table: 0 });
// @ts-expect-error A table is named by its position, not by text.
doc.deleteTable("Budget");
// @ts-expect-error The read session has no table methods.
reader.insertTable({ rows: 1, columns: 1 });
`;

/**
 * What an agent needs to name and place a table and no signature carries. Each phrase sits on
 * one line of the write declaration.
 */
const DOC_TABLE_GUIDANCE = [
  // A paragraph may read like a table; rows of a table inside a cell are not counted.
  "reading exactly `<table>`",
  "are indented further and are not counted",
  // Which context keeps `after` in place, and how a tab's first paragraph is named.
  "names a later paragraph",
  "the tab's first line included",
  "an empty cell before a filled one is `\"\"`",
  // Which tables take a column, and how far the tables of a tab grow.
  "only to a table of at most 500 rows",
  "only until the tables of one tab hold 5,000 cells",
  // A new cell takes its neighbour's look and none of what the HTML shows; which neighbour.
  "their bold, italic, strikethrough, links, list items or headings",
  "or of the last row when they are added at the end",
  "or of the last column when they are added at the end",
  // The refusals nothing in `getContent()` shows coming.
  "holds a dropdown or other smart chips and no text",
  "that holds nothing but a positioned image, or after a table whose next paragraph holds one",
  // `cells` goes by the table's rows and never sets how many columns are added.
  "For columns it is 1 unless given",
  "`cells` goes row by row",
  // What `replaceText()` does with table HTML, beside a table and inside one.
  "so is a table written as HTML",
  "unless the paragraph on that side of the table is read-only",
  "or, inside a table, the table's HTML",
  "Each parameter holds at most 2,000 lines",
  // Two tables that read the same.
  "make them differ first",
  "only while no other",
];

/** Guidance that every method it applies to repeats, as `[phrase, lines that hold it]`. */
const DOC_TABLE_REPEATED_GUIDANCE = [
  // A new table, new rows and new columns.
  ["at most 1,000 cells that have text", 3],
  // New rows and new columns.
  ["New cells hold plain paragraphs", 2],
  // The two methods that add rows and columns and the two that delete them.
  ["suggested changes to its rows, columns or cells", 4],
] as const;

function docBundle(): string {
  return [
    source("docs-read-types.txt"),
    stripTypeModulePrefix(source("docs-types.txt"), DOCS_TYPES_MODULE_PREFIX),
  ].join("\n");
}

function driveBundle(): string {
  return [
    source("docs-read-types.txt"),
    source("sheets-read-types.txt"),
    source("slides-read-types.txt"),
    stripTypeModulePrefix(source("drive-types.txt"), DRIVE_TYPES_MODULE_PREFIX),
  ].join("\n");
}

function sheetsBundle(): string {
  return [
    source("sheets-read-types.txt"),
    stripTypeModulePrefix(source("sheets-types.txt"), SHEETS_TYPES_MODULE_PREFIX),
  ].join("\n");
}

function slidesBundle(): string {
  return [
    source("slides-read-types.txt"),
    stripTypeModulePrefix(source("slides-types.txt"), SLIDES_TYPES_MODULE_PREFIX),
  ].join("\n");
}

/** The vendor's flat declaration bundle, composed as `GatekeeperVendor.getTypeScriptTypes()` does. */
function vendorBundle(): string {
  return [
    stripTypeModulePrefix(source("types.txt"), GMAIL_TYPES_MODULE_PREFIX), docBundle(),
    sheetsBundle(), slidesBundle(), source("calendar-types.txt"), source("bigquery-types.txt"),
    stripTypeModulePrefix(source("drive-types.txt"), DRIVE_TYPES_MODULE_PREFIX),
    source("chat-types.txt"),
  ].join("\n");
}

describe("embedded agent declarations", () => {
  it("exposes binary uploads and removable draft attachments without message write authority", () => {
    const types = source("types.txt");
    expect(compileAgentTypes(types + `
declare const draft: GmailDraft;
declare const message: GmailMessage;
const file: GmailAttachmentInput = {
  filename: "report.pdf", mimeType: "application/pdf", content: new ArrayBuffer(1),
};
const compose: GmailComposeOptions = {attachments: [file]};
const reply: GmailReplyOptions = {attachments: [file]};
const initial: GmailDraftInput = {attachments: [file]};
const added: Promise<GmailDraftAttachment> = draft.addAttachment(file);
const entries: Promise<GmailDraftAttachmentEntry[]> = draft.attachments();
declare const draftAttachment: GmailDraftAttachment;
const readable: GmailAttachment = draftAttachment;
const removed: Promise<void> = draftAttachment.remove();
// @ts-expect-error Patches cannot replace all attachments.
const patch: GmailDraftPatch = {attachments: [file]};
// @ts-expect-error Message attachment capabilities remain read-only.
readable.remove();
// @ts-expect-error Uploads take exact binary buffers, not base64 strings.
const encoded: GmailAttachmentInput = {filename: "file", mimeType: "text/plain", content: "YQ=="};
// @ts-expect-error Messages cannot add attachments.
message.addAttachment(file);
`).filter(message => !message.includes("'cloudflare:workers'"))).toEqual([]);
  });

  it("compiles the exact Google Doc agent declaration bundle without module dependencies", () => {
    expect(compileAgentTypes(docBundle())).toEqual([]);
  });

  it("compiles the exact Google Drive agent declaration bundle without module dependencies", () => {
    expect(compileAgentTypes(driveBundle())).toEqual([]);
  });

  it("compiles the exact Google Slides agent declaration bundle without module dependencies", () => {
    expect(compileAgentTypes(slidesBundle())).toEqual([]);
  });

  it("compiles the exact Google Sheets agent declaration bundle without module dependencies", () => {
    expect(compileAgentTypes(sheetsBundle())).toEqual([]);
  });

  it.each([
    ["vendor", vendorBundle, sheetsBundle],
    ["Drive", driveBundle, () => source("sheets-read-types.txt")],
  ] as const)(
    "declares each Google Sheets name once in the flat %s bundle",
    (_name, bundle, sheets) => {
      const sheetsNames = topLevelNames(sheets());
      expect(sheetsNames).toContain("GoogleSpreadsheetReadSession");
      expect(duplicates(topLevelNames(bundle())).filter(name => sheetsNames.includes(name)))
        .toEqual([]);
    },
  );

  it("declares the flattened tab contract on the canonical read session", () => {
    const readTypes = source("docs-read-types.d.ts");
    expect(readTypes).toContain("export type GoogleDocTab = {");
    expect(readTypes).toContain("listTabs(): Promise<GoogleDocTab[]>;");
    expect(readTypes).toContain("getContent(tabId?: string): Promise<string>;");
  });

  it.each([["Doc", docBundle], ["Drive", driveBundle]] as const)(
    "carries the exact tab shape and its selection guidance into the %s bundle",
    (_name, bundle) => {
      const types = bundle();
      for (const phrase of TAB_GUIDANCE) expect(types).toContain(phrase);
      expect(compileAgentTypes(types + TAB_SHAPE_CHECK)).toEqual([]);
    },
  );

  it("keeps Drive Docs authority read-only", () => {
    const readTypes = source("docs-read-types.d.ts");
    expect(readTypes).toContain("export interface GoogleDocReadSession");
    expect(readTypes).not.toContain("replaceText");
    expect(readTypes).not.toContain("appendText");
    for (const [name] of DOC_TABLE_METHODS) expect(readTypes).not.toContain(name);
    const writeTypes = source("docs-types.d.ts");
    expect(writeTypes).toContain(
      "export interface GoogleDocSession extends GoogleDocReadSession",
    );
    expect(writeTypes).toContain(
      "replaceText(oldMarkdown: string, newMarkdown: string, tabId?: string): Promise<void>;",
    );
    expect(writeTypes).toContain("appendText(markdown: string, tabId?: string): Promise<void>;");
    for (const [, signature] of DOC_TABLE_METHODS) expect(writeTypes).toContain(signature);
  });

  it("types table changes on a directly bound Doc", () => {
    expect(compileAgentTypes(docBundle() + DOC_TABLE_USAGE)).toEqual([]);
    // The limits are this connection's own, and the declaration gives no other reason for them.
    const writeTypes = source("docs-types.d.ts");
    for (const phrase of ["exempt", "Docs editor", "Google Docs allows"]) {
      expect(writeTypes).not.toContain(phrase);
    }
  });

  it("gives a Doc opened from Drive no table types", () => {
    const driveNames = topLevelNames(driveBundle());
    expect(driveNames).toContain("GoogleDocReadSession");
    expect(driveNames).not.toContain("GoogleDocSession");
    for (const name of DOC_TABLE_TYPE_NAMES) expect(driveNames).not.toContain(name);
  });

  it("declares each Doc table name once in the flat vendor bundle", () => {
    const names = topLevelNames(vendorBundle());
    // Slides declares table types of its own in the same flat scope.
    expect(names).toContain("TableElement");
    for (const name of DOC_TABLE_TYPE_NAMES) {
      expect(names.filter(candidate => candidate === name)).toEqual([name]);
    }
  });

  it("keeps what an agent needs to name and place a table", () => {
    const writeLines = source("docs-types.d.ts").split("\n");
    const linesHolding = (phrase: string) => writeLines.filter(line => line.includes(phrase));
    for (const phrase of DOC_TABLE_GUIDANCE) expect(linesHolding(phrase), phrase).not.toEqual([]);
    for (const [phrase, lineCount] of DOC_TABLE_REPEATED_GUIDANCE) {
      expect(linesHolding(phrase), phrase).toHaveLength(lineCount);
    }
    // How a merged cell is recognised, and that a list item's bullet character is not its text.
    const readTypes = source("docs-read-types.d.ts");
    expect(readTypes).toContain("A merged cell carries");
    expect(readTypes).toContain("starts with its bullet character");
  });

  it("shows a table as the renderer lays it out", () => {
    // An agent copies line breaks and indentation from the example in the getContent() comment
    // when a replacement spans cells, so the example is held to what the renderer writes.
    const commentPrefix = "   * ";
    const fence = commentPrefix + "```";
    const lines = source("docs-read-types.d.ts").split("\n");
    expect(lines.filter(line => line === fence)).toHaveLength(2);
    const exampleLines = lines.slice(lines.indexOf(fence) + 1, lines.lastIndexOf(fence));
    for (const line of exampleLines) expect(line.startsWith(commentPrefix), line).toBe(true);
    const example = exampleLines.map(line => line.slice(commentPrefix.length)).join("\n");

    const { markdown } = docTabToMarkdown(buildTab([
      { runs: ["Before\n"] },
      { table: [
        ["Name\n", "\n"],
        [
          { paragraphs: [{ runs: ["First paragraph\n"] }, { runs: ["Second paragraph\n"] }] },
          "Total\n",
        ],
      ] },
      { runs: ["After\n"] },
    ]));
    const tableEnd = "</table>";
    expect(markdown).toContain(tableEnd);
    expect(example).toBe(
      markdown.slice(markdown.indexOf("<table>"), markdown.indexOf(tableEnd) + tableEnd.length),
    );
  });

  it("gives a directly bound spreadsheet its own session and Drive only the read session", () => {
    expect(topLevelNames(vendorBundle())).toContain("GoogleSpreadsheetSession");
    expect(compileAgentTypes(vendorBundle() + `
declare const sheet: GoogleSpreadsheetSession;
const queued: Promise<void> = sheet.updateSpreadsheet([
  { op: "writeCells", range: "'Sales 2026'!A1:B1", values: [["=SUM(C1:C9)", null]] },
  { op: "clearRange", range: "Sales!A2:B3" },
]);
// @ts-expect-error A change names its range.
sheet.updateSpreadsheet([{ op: "clearRange" }]);
sheet.updateSpreadsheet([
  { op: "addSheet", title: "Q4", index: 1, rowCount: 100, columnCount: 5 },
  { op: "writeCells", range: "'Q4'!A1", values: [["Forecast"]] },
  { op: "renameSheet", sheet: "Q4", title: "Q4 2026" },
  { op: "duplicateSheet", sheet: "Sales", title: "Copy", index: 2 },
  { op: "insertRows", sheet: "Copy", at: 3, count: 2 },
  { op: "deleteRows", sheet: "Sales", at: 5 },
  { op: "insertColumns", sheet: "Sales", at: "C" },
  { op: "deleteColumns", sheet: "Sales", at: "C", count: 2 },
  { op: "deleteSheet", sheet: "Copy" },
]);
// @ts-expect-error Columns are named by their letters.
sheet.updateSpreadsheet([{ op: "insertColumns", sheet: "Sales", at: 3 }]);
// @ts-expect-error A sheet is named by its title, not its ID.
sheet.updateSpreadsheet([{ op: "deleteSheet", sheet: 0 }]);
// @ts-expect-error A sheet added needs a title.
sheet.updateSpreadsheet([{ op: "addSheet" }]);
`)
      // The agent's runtime provides the module Chat's declarations import.
      .filter(message => !message.includes("'cloudflare:workers'"))).toEqual([]);
    expect(compileAgentTypes(sheetsBundle() +
      "\nconst readable: GoogleSpreadsheetReadSession = null! as GoogleSpreadsheetSession;\n"))
      .toEqual([]);
    const driveNames = topLevelNames(driveBundle());
    expect(driveNames).toContain("GoogleSpreadsheetReadSession");
    expect(driveNames).not.toContain("GoogleSpreadsheetSession");
    expect(driveNames).not.toContain("SheetChange");
    expect(source("sheets-types.d.ts")).toContain(
      "export interface GoogleSpreadsheetSession extends GoogleSpreadsheetReadSession",
    );
  });

  it.each([
    ["vendor", vendorBundle],
    ["Drive", driveBundle],
  ] as const)("declares the format read types once in the flat %s bundle", (_name, bundle) => {
    const names = topLevelNames(bundle());
    const formatNames = [
      "SheetColor", "SheetBorderStyle", "SheetBorder", "SheetNumberFormat", "SheetCellFormat",
      "SpreadsheetFormats",
    ];
    // Slides declares colours, borders and text formats of its own in the same flat scope.
    expect(names).toContain("SlideColor");
    for (const name of formatNames) expect(names.filter(candidate => candidate === name)).toEqual([name]);
    expect(compileAgentTypes(bundle() + `
declare const sheet: GoogleSpreadsheetReadSession;
const read: Promise<SpreadsheetFormats> = sheet.readFormats("'Sales 2026'!A1:C3");
const format: SheetCellFormat | null = null! as SpreadsheetFormats["formats"][number][number];
const border: SheetBorder = { style: "SOLID_MEDIUM", color: "#ff0000" };
const set: SheetCellFormat = {
  bold: true, fontSize: 14, textColor: "ACCENT1", fillColor: "#ffffff", wrap: "CLIP",
  numberFormat: { type: "CURRENCY", pattern: '"$"#,##0.00' }, borders: { top: border },
};
// @ts-expect-error A border has a style.
const unstyled: SheetBorder = { color: "#000000" };
// @ts-expect-error Wrapping is named as the agent names it.
const legacy: SheetCellFormat = { wrap: "OVERFLOW_CELL" };
`).filter(message => !message.includes("'cloudflare:workers'"))).toEqual([]);
  });

  it("hands out only read-only native sessions from Drive", () => {
    const driveTypes = source("drive-types.d.ts");
    expect(driveTypes).toContain(
      "openGoogleDoc(fileId: string): Promise<GoogleDocReadSession>",
    );
    expect(driveTypes).toContain(
      "openGoogleSheet(fileId: string): Promise<GoogleSpreadsheetReadSession>",
    );
    expect(driveTypes).toContain(
      "openGoogleSlides(fileId: string): Promise<GooglePresentationReadSession>",
    );
    expect(driveTypes).not.toContain("GoogleDocSession>");
    expect(driveTypes).not.toContain("GoogleSpreadsheetSession>");
    expect(driveTypes).not.toContain("GooglePresentationSession>");
    expect(driveTypes).toContain("export interface GoogleDriveReadSession");
  });

  it("streams PDF text from both Drive sessions", () => {
    const driveTypes = source("drive-types.d.ts");
    expect(driveTypes).toContain("export type DrivePdfText = {");
    expect(driveTypes.split("readPdf(fileId: string): Promise<DrivePdfText>;")).toHaveLength(3);
    expect(compileAgentTypes(driveBundle() + `
declare const drive: GoogleDriveReadSession;
declare const folder: GoogleDriveFolderSession;
async function read(): Promise<string> {
  const { text, hasText } = await drive.readPdf("file-1");
  const flag: boolean = hasText;
  const reader = (await folder.readPdf("file-2")).text.getReader();
  const first: Uint8Array | undefined = (await reader.read()).value;
  await reader.cancel();
  return new Response(text).text();
}
// @ts-expect-error The text is a stream, not a string.
const whole: DrivePdfText = { text: "", hasText: true };
// @ts-expect-error Nothing is cut, so there is no truncation flag.
const cut: DrivePdfText = { text: new ReadableStream(), hasText: true, truncated: false };
`)).toEqual([]);
  });

  it("streams a file's bytes from both Drive sessions, whole or by range", () => {
    expect(compileAgentTypes(driveBundle() + `
declare const drive: GoogleDriveReadSession;
declare const folder: GoogleDriveFolderSession;
declare const worktree: { writeBytes(path: string, content: ReadableStream<Uint8Array>): Promise<void> };
async function readCsv(): Promise<string> {
  const whole = await drive.readFile("file-1");
  const text = await new Response(whole.content).text();
  const first = await folder.readFile("file-2", { offset: 16, length: 1024 });
  const rest: boolean = first.offset + first.length < first.size;
  const changed: boolean = first.modifiedTime.getTime() > 0;
  await worktree.writeBytes("data/file-2.bin", first.content);
  return text + first.mimeType + rest + changed;
}
// @ts-expect-error The content is a stream, not the bytes themselves.
const held: DriveFileBytes = { mimeType: "text/csv", size: 1, modifiedTime: new Date(), offset: 0, length: 1, content: new Uint8Array(1) };
`)).toEqual([]);
  });
});
