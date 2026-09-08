import { expect, it } from 'vitest';
import * as Y from 'yjs';

import { orderedYMaps } from './collaboration-order';
import { readTableData, reconcileTableData, writeTableData } from './collaboration-table';

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
	data.rows.unshift({ cells: [{ text: 'new' }] });
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
