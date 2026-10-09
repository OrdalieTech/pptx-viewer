import type { PptxElement } from 'pptx-viewer-core';
import { expect, test } from 'vitest';
import * as Y from 'yjs';

import { applyLivePatch } from './collaboration-live-patch-target';
import { reconcileSlidesInYDoc } from './collaboration-reconcile';
import { readSlidesFromYDoc, writeSlidesToYDoc } from './collaboration-sync';
import { mergeElement } from './element-operations';

const factories = {
	createMap: () => new Y.Map(),
	createArray: () => new Y.Array(),
	createText: () => new Y.Text(),
};
const table = (): PptxElement => ({
	type: 'table',
	id: 't',
	x: 0,
	y: 0,
	width: 200,
	height: 90,
	tableData: {
		columnWidths: [1],
		rows: [30, 60].map((height, i) => ({ height, cells: [{ text: String(i) }] })),
	},
});

test('scales only a geometry resize, preserving proportions, cells and explicit row changes', () => {
	const before = table();
	const next = mergeElement(before, { height: 180 });
	if (before.type !== 'table' || next.type !== 'table') {
		throw new Error('Expected tables');
	}
	expect(next.tableData!.rows.map((r) => r.height)).toStrictEqual([60, 120]);
	expect(before.tableData!.rows.map((r) => r.height)).toStrictEqual([30, 60]);
	expect(next.tableData!.rows[0].cells).toBe(before.tableData!.rows[0].cells);
	expect((mergeElement(before, { x: 40 }) as typeof next).tableData).toBe(before.tableData);
	expect((mergeElement(next, { height: 180 }) as typeof next).tableData).toBe(next.tableData);
	const explicit = { ...before.tableData!, rows: [{ height: 150, cells: [{ text: 'edited' }] }] };
	expect(
		(mergeElement(before, { height: 150, tableData: explicit }) as typeof next).tableData,
	).toBe(explicit);
});

test.each(['commit', 'live'] as const)(
	'keeps repeated %s resizes after a row edit across peers and reload',
	(mode) => {
		const doc = new Y.Doc();
		writeSlidesToYDoc(
			[{ id: 's', rId: 'r1', slideNumber: 1, elements: [table()] }],
			doc,
			factories,
		);
		const slides = readSlidesFromYDoc(doc);
		const element = slides[0].elements[0];
		if (element.type !== 'table') {
			throw new Error('Expected table');
		}
		element.tableData!.rows.push({ height: 30, cells: [{ text: 'added' }] });
		element.height = 120;
		reconcileSlidesInYDoc(slides, doc, factories);
		const peer = new Y.Doc();
		Y.applyUpdate(peer, Y.encodeStateAsUpdate(doc));
		for (const height of [240, 60, 180]) {
			const local = readSlidesFromYDoc(doc);
			local[0].elements[0] = mergeElement(local[0].elements[0], { height });
			if (mode === 'live') {
				applyLivePatch(
					doc,
					factories,
					{ slideId: 's', elementId: 't', geometry: { height } },
					100000,
				);
				expect(readSlidesFromYDoc(doc)[0].elements[0].height).toBe(height);
			}
			reconcileSlidesInYDoc(local, doc, factories);
			Y.applyUpdate(peer, Y.encodeStateAsUpdate(doc));
			const restored = readSlidesFromYDoc(peer)[0].elements[0];
			expect(restored.height).toBe(height);
			if (restored.type !== 'table') {
				throw new Error('Expected table');
			}
			expect(restored.tableData!.rows.map((r) => r.height)).toStrictEqual([
				height / 4,
				height / 2,
				height / 4,
			]);
			expect(restored.tableData!.rows.map((r) => r.cells[0].text)).toStrictEqual([
				'0',
				'1',
				'added',
			]);
		}
		doc.destroy();
		peer.destroy();
	},
);
