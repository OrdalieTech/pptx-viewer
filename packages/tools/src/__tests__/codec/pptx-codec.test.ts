import { ELEMENT_FIELD_KIND, PptxHandler, SLIDE_FIELD_KIND } from 'pptx-viewer-core';
import {
	ASSET_ELEMENT_FIELDS,
	COMPLEX_ELEMENT_FIELDS,
	COMPLEX_SLIDE_FIELDS,
} from 'pptx-viewer-shared/collaboration';
import { describe, it, expect, expectTypeOf } from 'vitest';
import { Doc as YDoc, encodeStateAsUpdate } from 'yjs';

import {
	COMPLEX_FIELD_MAP,
	COMPLEX_SLIDE_FIELD_MAP,
	ORIGIN_FILE_LOAD,
	PptxCodec,
	SCALAR_ELEMENT_KEYS,
	SCALAR_SLIDE_KEYS,
} from '../../codec/index.js';
import { createTestPptxBytes } from '../helpers/create-test-pptx.js';

describe('pptxCodec', () => {
	it('has correct formatId and extensions', () => {
		const codec = new PptxCodec();
		expect(codec.formatId).toBe('pptx');
		expect(codec.extensions).toContain('.pptx');
		expect(codec.extensions).not.toContain('.ppt');
	});

	it('exports ORIGIN_FILE_LOAD constant', () => {
		expect(ORIGIN_FILE_LOAD).toBe('file-load');
	});

	it('observe returns unsubscribe function', () => {
		const codec = new PptxCodec();
		const ydoc = new YDoc();
		let called = false;
		const unsub = codec.observe(ydoc, () => {
			called = true;
		});
		expectTypeOf(unsub).toBeFunction();

		// Trigger a change
		ydoc.getMap('pptx:meta').set('test', 'value');
		expect(called).toBeTruthy();

		// Unsubscribe
		unsub();
		called = false;
		ydoc.getMap('pptx:meta').set('test2', 'value2');
		// After unsubscribe, callback should not be called
		// (Yjs observe is synchronous, so this check is valid)
		expect(called).toBeFalsy();
	});
});

describe('pptxCodec hydrate', () => {
	it('hydrates a Y.Doc from real PPTX bytes', async () => {
		const codec = new PptxCodec();
		const ydoc = new YDoc();
		const bytes = await createTestPptxBytes(2);

		await codec.hydrate(ydoc, bytes);

		const meta = ydoc.getMap('pptx:meta');
		expect(meta.get('width')).toBeDefined();
		expect(meta.get('height')).toBeDefined();
		expect(meta.get('width')).toBeTypeOf('number');
		expect(meta.get('height')).toBeTypeOf('number');

		const slidesArray = ydoc.getArray('pptx:slides');
		expect(slidesArray).toHaveLength(2);
	});

	it('versions the schema without storing source bytes, including deleted Yjs history', async () => {
		const codec = new PptxCodec();
		const ydoc = new YDoc();
		const bytes = await createTestPptxBytes(1);

		await codec.hydrate(ydoc, bytes);

		const meta = ydoc.getMap('pptx:meta');
		expect(meta.has('sourceBytes')).toBeFalsy();
		expect(meta.get('schemaVersion')).toStrictEqual(expect.any(Number));
		expect(encodeStateAsUpdate(ydoc).length).toBeLessThan(bytes.length);
	});

	it('preserves slide data in Y.Doc', async () => {
		const codec = new PptxCodec();
		const ydoc = new YDoc();
		const bytes = await createTestPptxBytes(3);

		await codec.hydrate(ydoc, bytes);

		const slidesArray = ydoc.getArray('pptx:slides');
		expect(slidesArray).toHaveLength(3);

		// Each slide map should have an id
		for (let i = 0; i < slidesArray.length; i++) {
			const slideMap = slidesArray.get(i) as { get: (key: string) => unknown };
			expect(slideMap.get('id')).toBeTruthy();
		}
	});

	it('uses custom origin when provided', async () => {
		const codec = new PptxCodec();
		const ydoc = new YDoc();
		const bytes = await createTestPptxBytes(1);

		let capturedOrigin: unknown;
		ydoc.on('beforeTransaction', (tr: { origin: unknown }) => {
			capturedOrigin = tr.origin;
		});

		await codec.hydrate(ydoc, bytes, 'custom-origin');
		expect(capturedOrigin).toBe('custom-origin');
	});

	it('defaults to ORIGIN_FILE_LOAD origin', async () => {
		const codec = new PptxCodec();
		const ydoc = new YDoc();
		const bytes = await createTestPptxBytes(1);

		let capturedOrigin: unknown;
		ydoc.on('beforeTransaction', (tr: { origin: unknown }) => {
			capturedOrigin = tr.origin;
		});

		await codec.hydrate(ydoc, bytes);
		expect(capturedOrigin).toBe(ORIGIN_FILE_LOAD);
	});

	it('allows exactly one concurrent seed and preserves its snapshot', async () => {
		const codec = new PptxCodec();
		const ydoc = new YDoc();
		const sources = await Promise.all([createTestPptxBytes(1), createTestPptxBytes(3)]);
		const results = await Promise.allSettled(
			sources.map(async (source, index) => {
				await codec.hydrate(ydoc, source);
				return { snapshot: encodeStateAsUpdate(ydoc), slideCount: index === 0 ? 1 : 3 };
			}),
		);
		const successful = results.filter((result) => result.status === 'fulfilled');
		const rejected = results.filter((result) => result.status === 'rejected');
		expect(successful).toHaveLength(1);
		expect(rejected).toHaveLength(1);
		expect(rejected[0].reason.message).toMatch(/empty|existing|seed/i);
		expect(ydoc.getArray('pptx:slides')).toHaveLength(successful[0].value.slideCount);
		expect(encodeStateAsUpdate(ydoc)).toStrictEqual(successful[0].value.snapshot);
		ydoc.destroy();
	});

	it('rejects reseeding without replacing an existing collaborative snapshot', async () => {
		const codec = new PptxCodec();
		const ydoc = new YDoc();

		// First hydrate with 2 slides
		const bytes2 = await createTestPptxBytes(2);
		await codec.hydrate(ydoc, bytes2);
		expect(ydoc.getArray('pptx:slides')).toHaveLength(2);

		const snapshot = encodeStateAsUpdate(ydoc);
		// A different source must never replace existing collaborative state.
		const bytes3 = await createTestPptxBytes(3);
		await expect(codec.hydrate(ydoc, bytes3)).rejects.toThrow(/empty|existing|seed/i);
		expect(encodeStateAsUpdate(ydoc)).toStrictEqual(snapshot);
		expect(ydoc.getArray('pptx:slides')).toHaveLength(2);
	});
});

describe('pptxCodec dehydrate', () => {
	it('dehydrates Y.Doc back to PPTX bytes', async () => {
		const codec = new PptxCodec();
		const ydoc = new YDoc();
		const originalBytes = await createTestPptxBytes(2);

		await codec.hydrate(ydoc, originalBytes);
		const outputBytes = await codec.dehydrate(ydoc, originalBytes);

		expect(outputBytes).toBeInstanceOf(Uint8Array);
		expect(outputBytes.length).toBeGreaterThan(0);
	});

	it('dehydrated bytes produce valid PPTX', async () => {
		const codec = new PptxCodec();
		const ydoc = new YDoc();
		const originalBytes = await createTestPptxBytes(2);

		await codec.hydrate(ydoc, originalBytes);
		const outputBytes = await codec.dehydrate(ydoc, originalBytes);

		// Load the output to verify it's valid
		const handler = new PptxHandler();
		const pptxData = await handler.load(outputBytes.buffer as ArrayBuffer);
		expect(pptxData.slides).toHaveLength(2);
		expect(pptxData.width).toBeGreaterThan(0);
		expect(pptxData.height).toBeGreaterThan(0);
	});

	it('throws when no source bytes available', async () => {
		const codec = new PptxCodec();
		const ydoc = new YDoc();
		// Don't hydrate, just try to dehydrate an empty doc
		await expect(codec.dehydrate(ydoc, new Uint8Array())).rejects.toThrow();
	});
});

describe('pptxCodec fidelity', () => {
	it('round-trip preserves slide data deeply (load -> hydrate -> dehydrate -> reload)', async () => {
		const codec = new PptxCodec();
		const ydoc = new YDoc();
		const originalBytes = await createTestPptxBytes(3);

		const handler1 = new PptxHandler();
		const originalData = await handler1.load(originalBytes.buffer as ArrayBuffer);

		await codec.hydrate(ydoc, originalBytes);
		const outputBytes = await codec.dehydrate(ydoc, originalBytes);

		const handler2 = new PptxHandler();
		const roundTripData = await handler2.load(outputBytes.buffer as ArrayBuffer);

		expect(roundTripData.slides).toHaveLength(originalData.slides.length);
		for (let i = 0; i < originalData.slides.length; i++) {
			const orig = originalData.slides[i];
			const rt = roundTripData.slides[i];
			expect(rt.id).toBe(orig.id);
			expect(rt.elements).toHaveLength(orig.elements.length);
			for (let j = 0; j < orig.elements.length; j++) {
				const origEl = orig.elements[j];
				const rtEl = rt.elements[j];
				expect(rtEl.type).toBe(origEl.type);
				expect(rtEl.id).toBe(origEl.id);
				// Text segments must round-trip through Y.Text with the same
				// count and text content (style key order may differ).
				if ('textSegments' in origEl && origEl.textSegments) {
					expect(rtEl).toHaveProperty('textSegments');
					const origSegs = origEl.textSegments as Array<{
						text?: string;
						isParagraphBreak?: boolean;
					}>;
					const rtSegs = (rtEl as typeof origEl).textSegments as typeof origSegs;
					expect(rtSegs).toHaveLength(origSegs.length);
					const origText = origSegs.map((s) => s.text ?? '').join('');
					const rtText = rtSegs.map((s) => s.text ?? '').join('');
					expect(rtText).toBe(origText);
				}
			}
		}
	});

	it('textBody Y.Text is present in element Y.Map for text elements', async () => {
		const codec = new PptxCodec();
		const ydoc = new YDoc();
		const bytes = await createTestPptxBytes(1);
		await codec.hydrate(ydoc, bytes);

		const slidesArray = ydoc.getArray('pptx:slides');
		for (let i = 0; i < slidesArray.length; i++) {
			const slideMap = slidesArray.get(i) as { get: (k: string) => unknown };
			const elemArr = slideMap.get('elements') as
				| { length: number; get: (i: number) => unknown }
				| undefined;
			if (!elemArr) {
				continue;
			}
			for (let j = 0; j < elemArr.length; j++) {
				const elemMap = elemArr.get(j) as { get: (k: string) => unknown };
				const type = elemMap.get('type');
				// If the element has a textBody, verify it's a Y.Text (has toDelta)
				const textBody = elemMap.get('textBody');
				if (textBody) {
					expect((textBody as { toDelta?: unknown }).toDelta).toBeTypeOf('function');
					// _textSegments JSON blob should NOT exist (replaced by textBody)
					expect(elemMap.get('_textSegments')).toBeUndefined();
				}
				expect(type).toBeDefined();
			}
		}
	});
});

describe('pptxCodec round-trip', () => {
	it('preserves slide count through round-trip', async () => {
		const codec = new PptxCodec();
		const ydoc = new YDoc();
		const originalBytes = await createTestPptxBytes(3);

		// Load original
		const handler1 = new PptxHandler();
		const originalData = await handler1.load(originalBytes.buffer as ArrayBuffer);

		// Round-trip through codec
		await codec.hydrate(ydoc, originalBytes);
		const roundTrippedBytes = await codec.dehydrate(ydoc, originalBytes);

		const handler2 = new PptxHandler();
		const roundTrippedData = await handler2.load(roundTrippedBytes.buffer as ArrayBuffer);

		expect(roundTrippedData.slides).toHaveLength(originalData.slides.length);
	});

	it('preserves canvas dimensions through round-trip', async () => {
		const codec = new PptxCodec();
		const ydoc = new YDoc();
		const originalBytes = await createTestPptxBytes(1);

		const handler1 = new PptxHandler();
		const originalData = await handler1.load(originalBytes.buffer as ArrayBuffer);

		await codec.hydrate(ydoc, originalBytes);
		const roundTrippedBytes = await codec.dehydrate(ydoc, originalBytes);

		const handler2 = new PptxHandler();
		const roundTrippedData = await handler2.load(roundTrippedBytes.buffer as ArrayBuffer);

		expect(roundTrippedData.width).toBe(originalData.width);
		expect(roundTrippedData.height).toBe(originalData.height);
	});

	it('preserves element types through round-trip', async () => {
		const codec = new PptxCodec();
		const ydoc = new YDoc();
		const originalBytes = await createTestPptxBytes(2);

		const handler1 = new PptxHandler();
		const originalData = await handler1.load(originalBytes.buffer as ArrayBuffer);

		await codec.hydrate(ydoc, originalBytes);
		const roundTrippedBytes = await codec.dehydrate(ydoc, originalBytes);

		const handler2 = new PptxHandler();
		const roundTrippedData = await handler2.load(roundTrippedBytes.buffer as ArrayBuffer);

		for (let i = 0; i < originalData.slides.length; i++) {
			const origElements = originalData.slides[i].elements;
			const rtElements = roundTrippedData.slides[i].elements;
			expect(rtElements).toHaveLength(origElements.length);
			for (let j = 0; j < origElements.length; j++) {
				expect(rtElements[j].type).toBe(origElements[j].type);
			}
		}
	});

	it('observe fires callback on Y.Doc slide changes', async () => {
		const codec = new PptxCodec();
		const ydoc = new YDoc();
		const bytes = await createTestPptxBytes(1);
		await codec.hydrate(ydoc, bytes);

		let callCount = 0;
		const unsub = codec.observe(ydoc, () => {
			callCount++;
		});

		// Modify a slide in the Y.Doc
		const slidesArray = ydoc.getArray('pptx:slides');
		const slideMap = slidesArray.get(0) as { set: (k: string, v: unknown) => void };
		slideMap.set('notes', 'Updated notes');

		expect(callCount).toBeGreaterThan(0);
		unsub();
	});
});

describe('pptxCodec field-schema coverage', () => {
	it('scalar + complex element keys + textSegments cover every PptxElement field', () => {
		const coveredKind: Record<string, string> = { textSegments: 'text' };
		for (const key of SCALAR_ELEMENT_KEYS) {
			const kind = ELEMENT_FIELD_KIND[key as keyof typeof ELEMENT_FIELD_KIND];
			coveredKind[key] = kind === 'asset' ? 'asset' : 'scalar';
		}
		for (const key of Object.keys(COMPLEX_FIELD_MAP)) {
			coveredKind[key] = 'complex';
		}
		for (const key of ASSET_ELEMENT_FIELDS) {
			if (key in ELEMENT_FIELD_KIND) {
				coveredKind[key] = 'asset';
			}
		}
		// Group children are collaborative maps, not opaque JSON.
		coveredKind.children = 'nested';
		coveredKind.tableData = 'nested';

		for (const [field, kind] of Object.entries(ELEMENT_FIELD_KIND)) {
			expect(
				coveredKind[field],
				`field "${field}" is declared on PptxElement but not handled`,
			).toBe(kind);
		}
		expect(Object.keys(coveredKind).sort()).toStrictEqual(Object.keys(ELEMENT_FIELD_KIND).sort());
	});

	it('scalar + complex slide keys + elements cover every PptxSlide field', () => {
		const coveredKind: Record<string, string> = { elements: 'nested' };
		for (const key of SCALAR_SLIDE_KEYS) {
			coveredKind[key] = 'scalar';
		}
		for (const key of Object.keys(COMPLEX_SLIDE_FIELD_MAP)) {
			coveredKind[key] = 'complex';
		}

		for (const [field, kind] of Object.entries(SLIDE_FIELD_KIND)) {
			expect(coveredKind[field], `field "${field}" is declared on PptxSlide but not handled`).toBe(
				kind,
			);
		}
		expect(Object.keys(coveredKind).sort()).toStrictEqual(Object.keys(SLIDE_FIELD_KIND).sort());
	});

	it('no longer includes the removed phantom keys', () => {
		expect(SCALAR_ELEMENT_KEYS.has('placeholder')).toBeFalsy();
		expect(SCALAR_ELEMENT_KEYS.has('svgContent')).toBeFalsy();
		expect(SCALAR_ELEMENT_KEYS.has('inkSvg')).toBeFalsy();
		expect(SCALAR_ELEMENT_KEYS.has('sourceSlideId')).toBeFalsy();
		expect(COMPLEX_FIELD_MAP.connectionStart).toBeUndefined();
		expect(COMPLEX_FIELD_MAP.connectionEnd).toBeUndefined();
		expect(COMPLEX_FIELD_MAP.mediaBookmarks).toBeUndefined();
		expect(COMPLEX_FIELD_MAP).toStrictEqual(COMPLEX_ELEMENT_FIELDS);
		expect(COMPLEX_SLIDE_FIELD_MAP).toStrictEqual(COMPLEX_SLIDE_FIELDS);
	});

	it('covers newly-added binary/OLE/ink/3D fields', () => {
		expect(ASSET_ELEMENT_FIELDS.has('oleEmbeddedData')).toBeTruthy();
		expect(ASSET_ELEMENT_FIELDS.has('mediaData')).toBeTruthy();
		expect(ASSET_ELEMENT_FIELDS.has('modelData')).toBeTruthy();
		expect(SCALAR_ELEMENT_KEYS.has('inkPaths')).toBeTruthy();
		expect(COMPLEX_FIELD_MAP.extensionXml).toBe(COMPLEX_ELEMENT_FIELDS.extensionXml);
		expect(COMPLEX_FIELD_MAP.groupFill).toBe(COMPLEX_ELEMENT_FIELDS.groupFill);
		expect(COMPLEX_SLIDE_FIELD_MAP.notesShapes).toBe(COMPLEX_SLIDE_FIELDS.notesShapes);
		expect(COMPLEX_SLIDE_FIELD_MAP.headerFooterFlags).toBe(COMPLEX_SLIDE_FIELDS.headerFooterFlags);
	});
});
