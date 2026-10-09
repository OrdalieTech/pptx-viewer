import { expect, it } from 'vitest';
import * as Y from 'yjs';

import { orderedYMaps } from './collaboration-order';
import { readElementFromYMap } from './collaboration-sync';
import { readTableData, reconcileTableData, writeTableData } from './collaboration-table';
import { resolveTableCell } from './collaboration-table-columns';
import { insertTableColumn } from './table-layout';

const factories = {
	createMap: () => new Y.Map(),
	createArray: () => new Y.Array(),
	createText: () => new Y.Text(),
};

it('keeps cell identity and style across row/cell insertion and reorder, including concurrent text', () => {
	const doc = new Y.Doc();
	const element = doc.getMap('element');
	writeTableData(
		{ columnWidths: [100], rows: [{ cells: [{ text: 'original', style: { bold: true } }] }] },
		element,
		factories,
		'table',
	);
	const table = element.get('tableData') as Y.Map<unknown>;
	const rows = table.get('rows') as Y.Array<Y.Map<unknown>>;
	const originalRow = rows.get(0);
	const originalCell = (originalRow.get('cells') as Y.Array<Y.Map<unknown>>).get(0);
	const originalText = originalCell.get('textBody');
	const peer = new Y.Doc();
	Y.applyUpdate(peer, Y.encodeStateAsUpdate(doc));
	const peerRows = (peer.getMap('element').get('tableData') as Y.Map<unknown>).get(
		'rows',
	) as Y.Array<Y.Map<unknown>>;
	const peerText = (peerRows.get(0).get('cells') as Y.Array<Y.Map<unknown>>)
		.get(0)
		.get('textBody') as Y.Text;
	peerText.insert(8, '!');
	const data = readTableData(element)!;
	data.columnWidths = [50, 50];
	data.rows.unshift({ cells: [{ text: 'new' }, { text: '' }] });
	data.rows[1].cells.unshift({ text: 'new column' });
	reconcileTableData(data, element, factories, 'table');
	expect(orderedYMaps(rows)[1]).toBe(originalRow);
	expect(orderedYMaps(originalRow.get('cells') as Y.Array<Y.Map<unknown>>)[1]).toBe(originalCell);
	expect(originalCell.get('textBody')).toBe(originalText);
	Y.applyUpdate(doc, Y.encodeStateAsUpdate(peer));
	Y.applyUpdate(peer, Y.encodeStateAsUpdate(doc));
	const result = readTableData(element)!;
	expect(result.rows[1].cells[1]).toMatchObject({
		text: 'original!',
		collaborationId: 'table:cell:0:0',
		style: { bold: true },
	});
	expect(readTableData(peer.getMap('element'))).toEqual(result);
	result.rows.reverse();
	reconcileTableData(result, element, factories, 'table');
	expect(orderedYMaps(rows)[0]).toBe(originalRow);
	expect(readTableData(element)!.rows[0].cells[1].text).toBe('original!');
});

it('does not infer identity from position when replacement rows lack IDs', () => {
	const element = new Y.Doc().getMap('element');
	const data = { columnWidths: [100], rows: [{ cells: [{ text: 'same' }] }] };
	writeTableData(data, element, factories, 'table');
	const before = readTableData(element)!;
	reconcileTableData(data, element, factories, 'table');
	const after = readTableData(element)!;
	expect(after.rows[0].collaborationId).not.toBe(before.rows[0].collaborationId);
	expect(after.rows[0].cells[0].collaborationId).not.toBe(before.rows[0].cells[0].collaborationId);
});

it.each([0, 1, 2])('keeps simultaneous column insertions aligned at index %i', (index) => {
	const doc = new Y.Doc();
	const element = doc.getMap('element');
	writeTableData(
		{
			columnWidths: [0.5, 0.5],
			rows: [{ cells: [{ text: 'A' }, { text: 'B' }] }, { cells: [{ text: 'C' }, { text: 'D' }] }],
		},
		element,
		factories,
		'table',
	);
	const peer = new Y.Doc();
	Y.applyUpdate(peer, Y.encodeStateAsUpdate(doc));
	for (const [replica, label] of [
		[doc, 'left'],
		[peer, 'right'],
	] as const) {
		const map = replica.getMap('element');
		const table = readTableData(map)!;
		table.columnWidths[Math.min(index, 1)] /= 2;
		table.columnWidths.splice(index, 0, 0.25);
		table.rows.forEach((row, rowIndex) =>
			row.cells.splice(index, 0, { text: `${label}${rowIndex}` }),
		);
		reconcileTableData(table, map, factories, 'table');
	}
	Y.applyUpdate(doc, Y.encodeStateAsUpdate(peer));
	Y.applyUpdate(peer, Y.encodeStateAsUpdate(doc));
	const merged = readTableData(element)!;
	expect(readTableData(peer.getMap('element'))).toEqual(merged);
	expect(merged.columnWidths).toHaveLength(4);
	expect(merged.columnWidths.reduce((sum, width) => sum + width, 0)).toBeCloseTo(1);
	for (const label of ['left', 'right']) {
		const col = merged.rows[0].cells.findIndex((cell) => cell.text === `${label}0`);
		expect(merged.rows[1].cells[col].text).toBe(`${label}1`);
	}
	expect(
		merged.rows.map((row) =>
			row.cells.map((cell) => cell.text).filter((text) => text.length === 1),
		),
	).toEqual([
		['A', 'B'],
		['C', 'D'],
	]);
	merged.rows.forEach((row) => row.cells.splice(0, 1));
	merged.columnWidths.splice(0, 1);
	reconcileTableData(merged, element, factories, 'table');
	expect(readTableData(element)!.rows.every((row) => row.cells.length === 3)).toBe(true);
});

it('merges browser-normalized and codec-proportional column widths without mixing units', () => {
	const doc = new Y.Doc();
	const element = doc.getMap('element');
	writeTableData(
		{ columnWidths: [150, 150], rows: [{ cells: [{ text: 'A' }, { text: 'B' }] }] },
		element,
		factories,
		'table',
	);
	const peer = new Y.Doc();
	Y.applyUpdate(peer, Y.encodeStateAsUpdate(doc));
	const browser = insertTableColumn(readTableData(element)!, 0, 'left');
	reconcileTableData(browser, element, factories, 'table');
	const agent = readTableData(peer.getMap('element'))!;
	agent.columnWidths = [75, 75, 150];
	agent.rows[0].cells.unshift({ text: 'Agent' });
	reconcileTableData(agent, peer.getMap('element'), factories, 'table');
	Y.applyUpdate(doc, Y.encodeStateAsUpdate(peer));
	Y.applyUpdate(peer, Y.encodeStateAsUpdate(doc));
	const result = readTableData(element)!;
	const total = result.columnWidths.reduce((sum, width) => sum + width, 0);
	expect(result.columnWidths.map((width) => width / total).sort()).toEqual([0.2, 0.2, 0.2, 0.4]);
	expect(result.rows[0].cells.map((cell) => cell.text).sort()).toEqual(['', 'A', 'Agent', 'B']);
	expect(readTableData(peer.getMap('element'))).toEqual(result);
});

it('does not resurrect a deleted column when a concurrent insertion changes its width', () => {
	const doc = new Y.Doc();
	const element = doc.getMap('element');
	writeTableData(
		{ columnWidths: [0.5, 0.5], rows: [{ cells: [{ text: 'Delete me' }, { text: 'Keep me' }] }] },
		element,
		factories,
		'table',
	);
	reconcileTableData(readTableData(element)!, element, factories, 'table');
	const peer = new Y.Doc();
	Y.applyUpdate(peer, Y.encodeStateAsUpdate(doc));
	const inserted = insertTableColumn(readTableData(element)!, 0, 'left');
	inserted.rows[0].cells[0].text = 'Inserted';
	reconcileTableData(inserted, element, factories, 'table');
	const deleted = readTableData(peer.getMap('element'))!;
	deleted.columnWidths = [1];
	deleted.rows[0].cells.shift();
	reconcileTableData(deleted, peer.getMap('element'), factories, 'table');
	Y.applyUpdate(doc, Y.encodeStateAsUpdate(peer));
	Y.applyUpdate(peer, Y.encodeStateAsUpdate(doc));
	const result = readTableData(element)!;
	expect(result.rows[0].cells.map((cell) => cell.text)).toEqual(['Inserted', 'Keep me']);
	expect(result.columnWidths).toHaveLength(2);
	expect(readTableData(peer.getMap('element'))).toEqual(result);
});

it('keeps concurrent first writes to an intersection addressable after merging and structural edits', () => {
	const a = new Y.Doc();
	writeTableData(
		{ columnWidths: [100, 100], rows: [{ cells: [{ text: 'A' }, { text: 'B' }] }] },
		a.getMap('element'),
		factories,
		'table',
	);
	const b = new Y.Doc();
	Y.applyUpdate(b, Y.encodeStateAsUpdate(a));
	const addedRow = readTableData(a.getMap('element'))!;
	addedRow.rows.push({ cells: [{ text: '' }, { text: '' }] });
	reconcileTableData(addedRow, a.getMap('element'), factories, 'table');
	reconcileTableData(
		insertTableColumn(readTableData(b.getMap('element'))!, 0, 'right'),
		b.getMap('element'),
		factories,
		'table',
	);
	const au = Y.encodeStateAsUpdate(a),
		bu = Y.encodeStateAsUpdate(b);
	Y.applyUpdate(a, bu);
	Y.applyUpdate(b, au);
	const id = readTableData(a.getMap('element'))!.rows[1].cells[1].collaborationId!;
	for (const [doc, value] of [
		[a, 'Left'],
		[b, 'Right'],
	] as const) {
		const cell = resolveTableCell(doc.getMap('element'), id, factories)!;
		(cell.get('textBody') as Y.Text).insert(0, value);
	}
	const aFilled = Y.encodeStateAsUpdate(a),
		bFilled = Y.encodeStateAsUpdate(b);
	Y.applyUpdate(a, bFilled);
	Y.applyUpdate(b, aFilled);
	expect(readTableData(a.getMap('element'))).toEqual(readTableData(b.getMap('element')));
	for (const doc of [a, b]) {
		const cell = resolveTableCell(doc.getMap('element'), id, factories)!;
		expect((cell.get('textBody') as Y.Text).toString()).toBe(
			readTableData(doc.getMap('element'))!.rows[1].cells[1].text,
		);
		const data = readTableData(doc.getMap('element'))!;
		data.rows[1].cells[1].text = 'Final';
		reconcileTableData(data, doc.getMap('element'), factories, 'table');
		expect(readTableData(doc.getMap('element'))!.rows[1].cells[1]).toMatchObject({
			collaborationId: id,
			text: 'Final',
		});
	}
	a.destroy();
	b.destroy();
});

it('preserves imported frame height until rows change, then reads height from the current rows', () => {
	const doc = new Y.Doc();
	const element = doc.getMap('element');
	element.set('type', 'table');
	element.set('height', 100);
	const assets = doc.getMap('assets');
	writeTableData(
		{
			columnWidths: [100],
			rows: [
				{ height: 30, cells: [{ text: 'A' }] },
				{ height: 30, cells: [{ text: 'B' }] },
			],
		},
		element,
		factories,
		'table',
	);
	const height = () => readElementFromYMap(element, assets).height;
	expect(height()).toBe(100);
	const edited = readTableData(element)!;
	edited.rows[0].cells[0].text = 'Edited';
	reconcileTableData(edited, element, factories, 'table');
	expect(height()).toBe(100);
	const columnAdded = insertTableColumn(readTableData(element)!, 0, 'right');
	reconcileTableData(columnAdded, element, factories, 'table');
	expect(height()).toBe(100);
	const rowAdded = readTableData(element)!;
	rowAdded.rows.push({ height: 30, cells: [{ text: '' }, { text: '' }] });
	reconcileTableData(rowAdded, element, factories, 'table');
	expect(height()).toBe(90);
	const rowDeleted = readTableData(element)!;
	rowDeleted.rows.shift();
	reconcileTableData(rowDeleted, element, factories, 'table');
	expect(height()).toBe(60);
	const before = Y.encodeStateAsUpdate(doc);
	expect(height()).toBe(60);
	expect(Y.encodeStateAsUpdate(doc)).toEqual(before);
	doc.destroy();
});

it('retains a row and concurrent text when all rows are deleted by separate clients', () => {
	const doc = new Y.Doc();
	const element = doc.getMap('element');
	writeTableData(
		{
			columnWidths: [100],
			rows: [
				{ height: 30, cells: [{ text: 'A' }] },
				{ height: 30, cells: [{ text: 'B' }] },
			],
		},
		element,
		factories,
		'table',
	);
	const before = readTableData(element)!;
	const peer = new Y.Doc();
	Y.applyUpdate(peer, Y.encodeStateAsUpdate(doc));
	for (const [replica, index] of [
		[doc, 0],
		[peer, 1],
	] as const) {
		const map = replica.getMap('element');
		const data = readTableData(map)!;
		data.rows.splice(index, 1);
		reconcileTableData(data, map, factories, 'table');
		(
			resolveTableCell(map, data.rows[0].cells[0].collaborationId!, factories)!.get(
				'textBody',
			) as Y.Text
		).insert(1, '!');
	}
	const updates = [Y.encodeStateAsUpdate(doc), Y.encodeStateAsUpdate(peer)];
	Y.applyUpdate(doc, updates[1]);
	Y.applyUpdate(peer, updates[0]);
	const after = readTableData(element)!;
	expect(after.rows).toHaveLength(1);
	const retained = before.rows.find(
		(row) => row.collaborationId === after.rows[0].collaborationId,
	)!;
	expect(after.rows[0].cells[0]).toMatchObject({
		collaborationId: retained.cells[0].collaborationId,
		text: `${retained.cells[0].text}!`,
	});
	expect(readTableData(peer.getMap('element'))).toEqual(after);
	doc.destroy();
	peer.destroy();
});
