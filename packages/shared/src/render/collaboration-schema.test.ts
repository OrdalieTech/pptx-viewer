import type { PptxSlide } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';
import * as Y from 'yjs';

import { findElementYMap } from './collaboration-live-patch-target';
import { reconcileSlidesInYDoc } from './collaboration-reconcile';
import {
	readSlidesFromYDoc,
	writeSlidesToYDoc,
	registerCollaborationSource,
} from './collaboration-sync';
import type { YDocLike, YjsFactories } from './collaboration-sync';

const factories: YjsFactories = {
	createMap: () => new Y.Map(),
	createArray: () => new Y.Array(),
	createText: () => new Y.Text(),
};
const asDoc = (doc: Y.Doc) => doc as unknown as YDocLike;
const deck = (): PptxSlide[] => [
	{
		id: 's1',
		slideNumber: 1,
		elements: [
			{
				id: 'g',
				type: 'group',
				x: 0,
				y: 0,
				width: 100,
				height: 100,
				children: [
					{
						id: 't',
						type: 'text',
						x: 1,
						y: 2,
						width: 3,
						height: 4,
						text: 'hello',
						textSegments: [{ text: 'hello', style: {} }],
					},
				],
			},
		],
	},
	{ id: 's2', slideNumber: 2, elements: [] },
];

describe('versioned shared PPTX schema', () => {
	it('resolves the same native asset across data-URL and browser blob representations', () => {
		const source: PptxSlide[] = [
			{
				id: 'ppt/slides/slide1.xml',
				slideNumber: 1,
				elements: [
					{
						id: 'image',
						type: 'image',
						x: 0,
						y: 0,
						width: 1,
						height: 1,
						imagePath: 'ppt/media/image1.png',
						imageData: 'data:image/png;base64,AA==',
					},
				],
			},
		];
		const server = new Y.Doc();
		registerCollaborationSource(asDoc(server), source);
		writeSlidesToYDoc(source, asDoc(server), factories);
		const browser = new Y.Doc();
		Y.applyUpdate(browser, Y.encodeStateAsUpdate(server));
		const browserSource = structuredClone(source);
		(browserSource[0].elements[0] as { imageData: string }).imageData = 'blob:local-browser-object';
		registerCollaborationSource(asDoc(browser), browserSource);
		expect(
			(readSlidesFromYDoc(asDoc(browser))[0].elements[0] as { imageData: string }).imageData,
		).toBe('blob:local-browser-object');
		expect(browser.getMap('pptx:assets').size).toBe(0);
		server.destroy();
		browser.destroy();
	});

	it('merges a slide reorder with concurrent nested text and geometry edits', () => {
		const a = new Y.Doc(),
			b = new Y.Doc();
		writeSlidesToYDoc(deck(), asDoc(a), factories);
		Y.applyUpdate(b, Y.encodeStateAsUpdate(a));
		const first = a.getArray('pptx:slides').get(0);
		reconcileSlidesInYDoc(readSlidesFromYDoc(asDoc(a)).reverse(), asDoc(a), factories);
		expect(a.getArray('pptx:slides').get(0)).toBe(first);
		const text = findElementYMap(asDoc(b), 's1', 't')!;
		(text.get('textBody') as Y.Text).insert(5, '!');
		text.set('x', 42);
		Y.applyUpdate(a, Y.encodeStateAsUpdate(b));
		Y.applyUpdate(b, Y.encodeStateAsUpdate(a));
		expect(readSlidesFromYDoc(asDoc(a))).toStrictEqual(readSlidesFromYDoc(asDoc(b)));
		expect(readSlidesFromYDoc(asDoc(a)).map((s) => s.id)).toStrictEqual(['s2', 's1']);
		const restored = findElementYMap(asDoc(a), 's1', 't')!;
		expect((restored.get('textBody') as Y.Text).toString()).toBe('hello!');
		expect(restored.get('x')).toBe(42);
		a.destroy();
		b.destroy();
	});

	it('references existing image and background payloads without adding them to Yjs', () => {
		const image = `data:image/png;base64,${'A'.repeat(2_000_000)}`;
		const slides: PptxSlide[] = [
			{
				id: 's',
				slideNumber: 1,
				backgroundImage: image,
				elements: [{ id: 'i', type: 'image', x: 0, y: 0, width: 30, height: 30, imageData: image }],
			},
		];
		const a = new Y.Doc({ gc: false });
		registerCollaborationSource(asDoc(a), slides);
		writeSlidesToYDoc(slides, asDoc(a), factories);
		expect(Y.encodeStateAsUpdate(a).length).toBeLessThan(3000);
		expect(a.getMap('pptx:assets').size).toBe(0);
		expect(readSlidesFromYDoc(asDoc(a))[0].backgroundImage).toBe(image);
		const b = new Y.Doc();
		Y.applyUpdate(b, Y.encodeStateAsUpdate(a));
		registerCollaborationSource(asDoc(b), slides);
		expect(readSlidesFromYDoc(asDoc(b))).toStrictEqual(readSlidesFromYDoc(asDoc(a)));
		const next = readSlidesFromYDoc(asDoc(a));
		next[0].elements[0].x = 99;
		reconcileSlidesInYDoc(next, asDoc(a), factories);
		expect(Y.encodeStateAsUpdate(a).length).toBeLessThan(4000);
		a.destroy();
		b.destroy();
	});

	it('rejects unknown schema versions and legacy embedded binary states', () => {
		const doc = new Y.Doc();
		doc.getMap('pptx:meta').set('schemaVersion', 999);
		expect(() => readSlidesFromYDoc(asDoc(doc))).toThrow('Unsupported');
		doc.getMap('pptx:meta').set('schemaVersion', 1);
		doc.getMap('pptx:meta').set('sourceBytes', new Y.Array());
		expect(() => readSlidesFromYDoc(asDoc(doc))).toThrow('Legacy');
		doc.destroy();
	});
});
