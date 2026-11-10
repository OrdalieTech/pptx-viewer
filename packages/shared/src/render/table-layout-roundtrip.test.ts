/**
 * Row and column edits of a table with a rectangular merge must survive a
 * save and a reload.
 *
 * The unit tests of `table-layout` prove the model. They cannot prove that the
 * merge the model describes is one the save writer serialises and the loader
 * reads back with the same shape, nor that the anchor's rich text reaches
 * OOXML when another cell takes the merge over. This drives a deck through the
 * full circle for each edit.
 */
import { PptxHandler, PresentationBuilder } from 'pptx-viewer-core';
import type { PptxSlide, PptxTableData, TablePptxElement } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import {
	deleteTableColumn,
	deleteTableRow,
	insertTableColumn,
	insertTableRow,
} from './table-layout';

interface Loaded {
	handler: PptxHandler;
	slides: PptxSlide[];
	table: TablePptxElement;
}

async function load(bytes: Uint8Array): Promise<Loaded> {
	const handler = new PptxHandler();
	const { slides } = await handler.load(bytes.slice().buffer as ArrayBuffer);
	const table = slides[0].elements.find((element) => element.type === 'table');
	return { handler, slides, table: table as TablePptxElement };
}

/**
 * A 3x3 table whose top-left 2x2 block is one merged cell with two formatted
 * runs, centred, as a package (the merge is set on a loaded
 * table: the builder writes plain cells).
 */
async function mergedDeck(): Promise<Uint8Array> {
	const { handler, data, createSlide } = await PresentationBuilder.create();
	const rows = [0, 1, 2].map((r) => ({ cells: [0, 1, 2].map((c) => ({ text: `r${r}c${c}` })) }));
	data.slides.push(
		createSlide('Blank').addTable({ rows }, { x: 0, y: 0, width: 600, height: 200 }).build(),
	);
	const plain = await load(await handler.save(data.slides));
	const cells = plain.table.tableData!.rows.map((row) => row.cells);
	cells[0][0] = {
		...cells[0][0],
		text: 'Bold plain',
		style: { ...cells[0][0].style, align: 'center' },
		textRuns: undefined,
		textSegments: [
			{ text: 'Bold ', style: { bold: true, fontFamily: 'Georgia' } },
			{ text: 'plain', style: { italic: true } },
		],
		gridSpan: 2,
		rowSpan: 2,
	};
	cells[0][1] = {
		...cells[0][1],
		text: '',
		textRuns: undefined,
		textSegments: undefined,
		hMerge: true,
	};
	cells[1][0] = {
		...cells[1][0],
		text: '',
		textRuns: undefined,
		textSegments: undefined,
		vMerge: true,
	};
	cells[1][1] = {
		...cells[1][1],
		text: '',
		textRuns: undefined,
		textSegments: undefined,
		hMerge: true,
		vMerge: true,
	};
	return plain.handler.save(plain.slides);
}

/** Each cell as `<columns>x<rows>`, `h`/`v` for a continuation, then its text. */
const shape = (table: PptxTableData): string[] =>
	table.rows.map((row) =>
		row.cells
			.map(
				(cell) =>
					`${cell.gridSpan ?? 1}x${cell.rowSpan ?? 1}${cell.hMerge ? 'h' : ''}${cell.vMerge ? 'v' : ''}:${cell.text}`,
			)
			.join(' | '),
	);

/** The runs of a cell with the formatting the fixture sets. */
const rich = (table: PptxTableData, row: number, col: number): unknown[] =>
	(table.rows[row].cells[col].textSegments ?? [])
		.filter((segment) => segment.text)
		.map((segment) => [
			segment.text,
			segment.style.bold ?? false,
			segment.style.italic ?? false,
			segment.style.fontFamily,
			table.rows[row].cells[col].style?.align,
		]);

const ANCHOR_RUNS = [
	['Bold ', true, false, 'Georgia', 'center'],
	['plain', false, true, undefined, 'center'],
];

/** Apply `edit` to the merged table, save, and read the table back. */
async function roundTrip(edit: (table: PptxTableData) => PptxTableData): Promise<PptxTableData> {
	const loaded = await load(await mergedDeck());
	loaded.table.tableData = edit(loaded.table.tableData!);
	const reloaded = await load(await loaded.handler.save(loaded.slides));
	return reloaded.table.tableData!;
}

describe('table-layout edits of a rectangular merge survive save and reload', () => {
	it('the fixture loads as one 2x2 merge with its rich text', async () => {
		const table = await roundTrip((data) => data);
		expect(shape(table)).toStrictEqual([
			'2x2:Bold plain | 1x1h: | 1x1:r0c2',
			'1x1v: | 1x1hv: | 1x1:r1c2',
			'1x1:r2c0 | 1x1:r2c1 | 1x1:r2c2',
		]);
		expect(rich(table, 0, 0)).toStrictEqual(ANCHOR_RUNS);
	});

	it('a row inserted inside the merge is covered over the whole width of the merge', async () => {
		const table = await roundTrip((data) => insertTableRow(data, 0, 'below'));
		expect(shape(table)).toStrictEqual([
			'2x3:Bold plain | 1x1h: | 1x1:r0c2',
			'1x1v: | 1x1hv: | 1x1:',
			'1x1v: | 1x1hv: | 1x1:r1c2',
			'1x1:r2c0 | 1x1:r2c1 | 1x1:r2c2',
		]);
		expect(rich(table, 0, 0)).toStrictEqual(ANCHOR_RUNS);
	});

	it('a column inserted inside the merge is covered over the whole height of the merge', async () => {
		const table = await roundTrip((data) => insertTableColumn(data, 0, 'right'));
		expect(shape(table)).toStrictEqual([
			'3x2:Bold plain | 1x1h: | 1x1h: | 1x1:r0c2',
			'1x1v: | 1x1hv: | 1x1hv: | 1x1:r1c2',
			'1x1:r2c0 | 1x1: | 1x1:r2c1 | 1x1:r2c2',
		]);
		expect(rich(table, 0, 0)).toStrictEqual(ANCHOR_RUNS);
	});

	it('deleting the anchor row moves the merge and its rich text to the next row', async () => {
		const table = await roundTrip((data) => deleteTableRow(data, 0));
		expect(shape(table)).toStrictEqual([
			'2x1:Bold plain | 1x1h: | 1x1:r1c2',
			'1x1:r2c0 | 1x1:r2c1 | 1x1:r2c2',
		]);
		expect(rich(table, 0, 0)).toStrictEqual(ANCHOR_RUNS);
	});

	it('deleting the anchor column moves the merge and its rich text to the next column', async () => {
		const table = await roundTrip((data) => deleteTableColumn(data, 0));
		expect(shape(table)).toStrictEqual([
			'1x2:Bold plain | 1x1:r0c2',
			'1x1v: | 1x1:r1c2',
			'1x1:r2c1 | 1x1:r2c2',
		]);
		expect(rich(table, 0, 0)).toStrictEqual(ANCHOR_RUNS);
	});
});
