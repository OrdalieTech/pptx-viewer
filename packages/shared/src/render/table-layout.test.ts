import type { PptxTableData } from 'pptx-viewer-core';
import { describe, it, expect } from 'vitest';

import {
	insertTableRow,
	deleteTableRow,
	insertTableColumn,
	deleteTableColumn,
} from './table-layout';

function makeTable(rows: number, cols: number): PptxTableData {
	return {
		rows: Array.from({ length: rows }, () => ({
			cells: Array.from({ length: cols }, () => ({ text: '' })),
		})),
		columnWidths: Array.from({ length: cols }, () => 1 / cols),
	} as PptxTableData;
}

describe('insertTableRow', () => {
	it('should add a row below the reference index', () => {
		const table = makeTable(2, 3);
		const result = insertTableRow(table, 0, 'below');
		expect(result.rows).toHaveLength(3);
		expect(result.rows[1].cells).toHaveLength(3);
	});

	it('should add a row above the reference index', () => {
		const table = makeTable(2, 3);
		const result = insertTableRow(table, 1, 'above');
		expect(result.rows).toHaveLength(3);
	});

	it('should not mutate the original table', () => {
		const table = makeTable(2, 2);
		insertTableRow(table, 0, 'below');
		expect(table.rows).toHaveLength(2);
	});

	it('should grow a vertical merge anchor that straddles the insertion point', () => {
		const table = makeTable(3, 1);
		table.rows[0].cells[0].rowSpan = 3;
		table.rows[1].cells[0] = { text: '', vMerge: true };
		table.rows[2].cells[0] = { text: '', vMerge: true };
		const result = insertTableRow(table, 1, 'below');
		expect(result.rows[0].cells[0].rowSpan).toBe(4);
		// The inserted row's cell is a vMerge continuation.
		expect(result.rows[2].cells[0].vMerge).toBeTruthy();
	});
});

describe('deleteTableRow', () => {
	it('should remove the row at the index', () => {
		const table = makeTable(3, 2);
		const result = deleteTableRow(table, 1);
		expect(result.rows).toHaveLength(2);
	});

	it('should refuse to remove the only row', () => {
		const table = makeTable(1, 2);
		expect(deleteTableRow(table, 0)).toBe(table);
	});

	it('should return original for out-of-range index', () => {
		const table = makeTable(2, 2);
		expect(deleteTableRow(table, 5)).toBe(table);
	});

	it('should decrement the anchor rowSpan when deleting a continuation row', () => {
		const table = makeTable(3, 1);
		table.rows[0].cells[0].rowSpan = 3;
		table.rows[1].cells[0] = { text: '', vMerge: true };
		table.rows[2].cells[0] = { text: '', vMerge: true };
		const result = deleteTableRow(table, 2);
		expect(result.rows[0].cells[0].rowSpan).toBe(2);
	});
});

describe('insertTableColumn', () => {
	it('should add a column to the right and renormalise widths', () => {
		const table = makeTable(2, 2);
		const result = insertTableColumn(table, 0, 'right');
		expect(result.columnWidths).toHaveLength(3);
		const sum = result.columnWidths.reduce((a, b) => a + b, 0);
		expect(sum).toBeCloseTo(1, 5);
		expect(result.rows[0].cells).toHaveLength(3);
	});

	it('should add a column to the left', () => {
		const table = makeTable(1, 2);
		const result = insertTableColumn(table, 1, 'left');
		expect(result.rows[0].cells).toHaveLength(3);
	});

	it('should grow a horizontal merge anchor that straddles the insertion point', () => {
		const table = makeTable(1, 3);
		table.rows[0].cells[0].gridSpan = 3;
		table.rows[0].cells[1] = { text: '', hMerge: true };
		table.rows[0].cells[2] = { text: '', hMerge: true };
		const result = insertTableColumn(table, 1, 'right');
		expect(result.rows[0].cells[0].gridSpan).toBe(4);
	});
});

describe('deleteTableColumn', () => {
	it('should remove the column and renormalise widths', () => {
		const table = makeTable(2, 3);
		const result = deleteTableColumn(table, 1);
		expect(result.columnWidths).toHaveLength(2);
		const sum = result.columnWidths.reduce((a, b) => a + b, 0);
		expect(sum).toBeCloseTo(1, 5);
		expect(result.rows[0].cells).toHaveLength(2);
	});

	it('should refuse to remove the only column', () => {
		const table = makeTable(2, 1);
		expect(deleteTableColumn(table, 0)).toBe(table);
	});

	it('should return original for out-of-range index', () => {
		const table = makeTable(2, 2);
		expect(deleteTableColumn(table, 9)).toBe(table);
	});

	it('should decrement the anchor gridSpan when deleting a continuation column', () => {
		const table = makeTable(1, 3);
		table.rows[0].cells[0].gridSpan = 3;
		table.rows[0].cells[1] = { text: '', hMerge: true };
		table.rows[0].cells[2] = { text: '', hMerge: true };
		const result = deleteTableColumn(table, 2);
		expect(result.rows[0].cells[0].gridSpan).toBe(2);
	});
});

/** A 3x3 table whose top-left 2x2 block is one merge; `repeat` writes the spans on the continuation cells as PowerPoint does. */
function makeMerged(repeat = false): PptxTableData {
	const table = makeTable(3, 3);
	table.rows.forEach((row, r) => row.cells.forEach((cell, c) => (cell.text = `r${r}c${c}`)));
	table.rows[0].cells[0] = {
		text: 'anchor',
		gridSpan: 2,
		rowSpan: 2,
		style: { bold: true },
		textRuns: [{ text: 'anchor', bold: true }],
		textSegments: [{ text: 'anchor', style: { bold: true } }],
	};
	table.rows[0].cells[1] = { text: '', hMerge: true, ...(repeat ? { rowSpan: 2 } : {}) };
	table.rows[1].cells[0] = {
		text: '',
		vMerge: true,
		...(repeat ? { gridSpan: 2 } : {}),
		textRuns: [],
		textSegments: [],
	};
	table.rows[1].cells[1] = { text: '', hMerge: true, vMerge: true };
	return table;
}
const merges = (table: PptxTableData): string[] =>
	table.rows.map((row) =>
		row.cells
			.map(
				(cell) =>
					`${cell.gridSpan ?? 1}x${cell.rowSpan ?? 1}${cell.hMerge ? 'h' : ''}${cell.vMerge ? 'v' : ''}`,
			)
			.join(' '),
	);

describe('rectangular merges', () => {
	it.each([false, true])(
		'a row inserted inside covers every column of the merge (spans repeated: %s)',
		(repeat) => {
			expect(merges(insertTableRow(makeMerged(repeat), 0, 'below'))).toStrictEqual(
				repeat
					? ['2x3 1x3h 1x1', '2x1v 1x1hv 1x1', '2x1v 1x1hv 1x1', '1x1 1x1 1x1']
					: ['2x3 1x1h 1x1', '1x1v 1x1hv 1x1', '1x1v 1x1hv 1x1', '1x1 1x1 1x1'],
			);
		},
	);

	it.each([false, true])(
		'a column inserted inside covers every row of the merge (spans repeated: %s)',
		(repeat) => {
			expect(merges(insertTableColumn(makeMerged(repeat), 0, 'right'))).toStrictEqual(
				repeat
					? ['3x2 1x2h 1x2h 1x1', '3x1v 1x1hv 1x1hv 1x1', '1x1 1x1 1x1 1x1']
					: ['3x2 1x1h 1x1h 1x1', '1x1v 1x1hv 1x1hv 1x1', '1x1 1x1 1x1 1x1'],
			);
		},
	);

	it('a row or column inserted beside the merge stays independent', () => {
		expect(merges(insertTableRow(makeMerged(), 1, 'below'))[2]).toBe('1x1 1x1 1x1');
		expect(
			merges(insertTableColumn(makeMerged(), 1, 'right')).map((row) => row.split(' ')[2]),
		).toStrictEqual(['1x1', '1x1', '1x1']);
	});

	it.each([false, true])(
		'deleting the anchor row promotes the next row with the rich content (spans repeated: %s)',
		(repeat) => {
			const result = deleteTableRow(makeMerged(repeat), 0);
			expect(merges(result)).toStrictEqual(['2x1 1x1h 1x1', '1x1 1x1 1x1']);
			expect(result.rows[0].cells[0]).toMatchObject({
				text: 'anchor',
				style: { bold: true },
				textRuns: [{ text: 'anchor', bold: true }],
				textSegments: [{ text: 'anchor', style: { bold: true } }],
			});
		},
	);

	it.each([false, true])(
		'deleting the anchor column promotes the next column with the rich content (spans repeated: %s)',
		(repeat) => {
			const result = deleteTableColumn(makeMerged(repeat), 0);
			expect(merges(result)).toStrictEqual(['1x2 1x1', '1x1v 1x1', '1x1 1x1']);
			expect(result.rows[0].cells[0]).toMatchObject({
				text: 'anchor',
				style: { bold: true },
				textRuns: [{ text: 'anchor', bold: true }],
				textSegments: [{ text: 'anchor', style: { bold: true } }],
			});
		},
	);

	it('an empty anchor leaves the promoted cell its own text', () => {
		const table = makeMerged();
		table.rows[0].cells[0] = { text: '', gridSpan: 2, rowSpan: 2 };
		table.rows[1].cells[0] = { text: 'kept', vMerge: true };
		expect(deleteTableRow(table, 0).rows[0].cells[0]).toMatchObject({ text: 'kept', gridSpan: 2 });
	});
});
