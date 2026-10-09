import { PptxHandler } from 'pptx-viewer-core';
import type { TablePptxElement } from 'pptx-viewer-core';
import { readSlidesFromYDoc, reconcileSlidesInYDoc, writeSlidesToYDoc } from 'pptx-viewer-shared';
import { expect, test } from 'vitest';
import * as Y from 'yjs';

import { EditorState } from './editor-state.svelte';

const factories = {
	createMap: () => new Y.Map(),
	createArray: () => new Y.Array(),
	createText: () => new Y.Text(),
};

test.each(['insert', 'delete'])(
	'resizes after a row %s and preserves rows through collaboration and PPTX export',
	async (action) => {
		const { handler, data } = await PptxHandler.create({ initialSlideCount: 1 });
		const doc = new Y.Doc();
		const reopened = new PptxHandler();
		try {
			data.slides[0].elements = [
				{
					id: 't',
					type: 'table',
					x: 20,
					y: 30,
					width: 200,
					height: 90,
					tableData: {
						columnWidths: [1],
						rows: [30, 30, 30].map((height, i) => ({ height, cells: [{ text: String(i) }] })),
					},
				},
			];
			writeSlidesToYDoc(data.slides, doc, factories);
			const slides = readSlidesFromYDoc(doc);
			const table = slides[0].elements[0] as TablePptxElement;
			if (action === 'insert') {
				table.tableData!.rows.push({ height: 30, cells: [{ text: 'new' }] });
			} else {
				table.tableData!.rows.pop();
			}
			reconcileSlidesInYDoc(slides, doc, factories);
			const editor = new EditorState({
				getCurrent: () => 0,
				getHandler: () => handler,
				onChange: () => {},
			});
			editor.editable = true;
			editor.setSlides(readSlidesFromYDoc(doc));
			const texts = (editor.slides[0].elements[0] as TablePptxElement).tableData!.rows.map(
				(r) => r.cells[0].text,
			);
			for (const height of [180, 120]) {
				editor.patchGeometry('t', { x: 20, y: 30, width: 200, height, rotation: 0 });
				reconcileSlidesInYDoc(editor.slides, doc, factories);
				const result = readSlidesFromYDoc(doc)[0].elements[0] as TablePptxElement;
				expect(result.height).toBe(height);
				expect(result.tableData!.rows.map((r) => r.height)).toStrictEqual(
					texts.map(() => height / texts.length),
				);
			}
			const bytes = await handler.save(readSlidesFromYDoc(doc));
			const result = (await reopened.load(bytes)).slides[0].elements.find(
				(e) => e.type === 'table',
			) as TablePptxElement;
			expect(result.height).toBeCloseTo(120);
			expect(result.tableData!.rows.map((r) => r.cells[0].text)).toStrictEqual(texts);
			expect(result.tableData!.rows.map((r) => r.height)).toStrictEqual(
				texts.map(() => 120 / texts.length),
			);
		} finally {
			doc.destroy();
			handler.dispose();
			reopened.dispose();
		}
	},
);
