import JSZip from 'jszip';
import {
	readSlidesFromYDoc,
	registerCollaborationSource,
	writeSlidesToYDoc,
} from 'pptx-viewer-shared/collaboration';
import { describe, expect, it } from 'vitest';
import {
	Array as YArray,
	Doc,
	Map as YMap,
	Text,
	encodeStateAsUpdate,
	encodeStateVector,
} from 'yjs';

import { PptxCodec } from '../../codec/index.js';
import { createTestPptxBytes } from '../helpers/create-test-pptx.js';
import {
	elementMaps,
	expectPackagePreserved,
	factories,
	fixtureBytes,
	fixtureMissing,
	flattenElements,
	loadPptx,
	replaceText,
	restoredDoc,
	slideMaps,
	xmlParser,
} from './fidelity-helpers.js';

const mediaFixture = 'Image_JPG_PNG_Audio_M4_A_Video_MP_4_12_Slides_36_8_MB_ff1095731b.pptx';
const cases = [
	{ fixture: 'sample-deck.pptx', index: 4, feature: 'table' },
	{ fixture: 'chart-gallery.pptx', index: 0, feature: 'chart' },
	{ fixture: 'master-views.pptx', index: 0, feature: 'master/layout' },
	{ fixture: 'transitions-animations.pptx', index: 2, feature: 'animation' },
	{ fixture: 'ole-embed.pptx', index: 0, feature: 'ole' },
	{ fixture: 'absolute-path-rels.pptx', index: 0, feature: 'group' },
	{ fixture: mediaFixture, index: 10, feature: 'media' },
];

describe('real PPTX package fidelity', () => {
	for (const { fixture, index, feature } of cases) {
		it.skipIf(fixtureMissing(fixture))(
			`${feature}: a dirty slide edit preserves other elements, parts and relationships (${fixture})`,
			async () => {
				const source = await fixtureBytes(fixture);
				const original = await loadPptx(source);
				const codec = new PptxCodec();
				const doc = new Doc();
				await codec.hydrate(doc, source);
				const slide = slideMaps(doc)[index];
				expect(slide).toBeDefined();
				// Prevent a false pass through core's unchanged-slide fast path.
				for (const map of slideMaps(doc)) {
					map.set('isDirty', false);
				}
				const before = original.slides[index];
				if (['chart', 'table', 'media', 'ole', 'group'].includes(feature)) {
					expect(before.elements.some((element) => element.type === feature)).toBeTruthy();
				}
				if (feature === 'animation') {
					expect(before.rawTiming).toBeDefined();
				}
				if (feature === 'master/layout') {
					const zip = await JSZip.loadAsync(source);
					expect(zip.file(/^ppt\/slideMasters\/.*\.xml$/).length).toBeGreaterThan(0);
					expect(zip.file(/^ppt\/slideLayouts\/.*\.xml$/).length).toBeGreaterThan(0);
				}
				let editedId: string;
				if (feature === 'ole' || feature === 'media') {
					const element = elementMaps(slide)[0];
					editedId = element.get('id') as string;
					element.set('x', Number(element.get('x')) + 12);
					slide.set('isDirty', true);
				} else {
					editedId = replaceText(slide, 'Collaborative fidelity edit');
				}
				expect(slide.get('isDirty')).toBeTruthy();
				const persisted = restoredDoc(doc);
				const output = await new PptxCodec().dehydrate(persisted, source);
				const actual = await loadPptx(output);
				expect(actual.slides.map((item) => item.id)).toStrictEqual(
					original.slides.map((item) => item.id),
				);
				const after = actual.slides[index];
				expect(flattenElements(after.elements).map((item) => [item.id, item.type])).toStrictEqual(
					flattenElements(before.elements).map((item) => [item.id, item.type]),
				);
				for (const element of flattenElements(before.elements)) {
					const reloaded = flattenElements(after.elements).find((item) => item.id === element.id)!;
					if (element.id === editedId) {
						if (feature === 'ole' || feature === 'media') {
							expect(reloaded.x).toBeCloseTo(element.x + 12, 3);
						} else {
							expect('text' in reloaded && reloaded.text).toBe('Collaborative fidelity edit');
						}
					} else {
						for (const key of [
							'x',
							'y',
							'width',
							'height',
							'rotation',
							'text',
							'textSegments',
							'textStyle',
							'shapeStyle',
							'chartData',
							'tableData',
							'cropLeft',
							'cropRight',
							'cropTop',
							'cropBottom',
							'imageData',
						]) {
							if (key === 'textSegments' && Reflect.get(element, key)) {
								// Explicit defaults may be materialized; existing run content/styles must survive.
								expect(Reflect.get(reloaded, key), `${fixture}:${element.id}:${key}`).toMatchObject(
									Reflect.get(element, key),
								);
							} else {
								expect(Reflect.get(reloaded, key), `${fixture}:${element.id}:${key}`).toStrictEqual(
									Reflect.get(element, key),
								);
							}
						}
					}
				}
				const originalZip = await JSZip.loadAsync(source);
				const outputZip = await JSZip.loadAsync(output);
				const originalXml = xmlParser.parse(await originalZip.file(before.id)!.async('string'))[
					'p:sld'
				];
				const outputXml = xmlParser.parse(await outputZip.file(before.id)!.async('string'))[
					'p:sld'
				];
				for (const key of ['p:timing', 'p:transition']) {
					expect(outputXml[key], `${fixture}:${key}`).toStrictEqual(originalXml[key]);
				}
				await expectPackagePreserved(source, output, [before.id]);
				doc.destroy();
				persisted.destroy();
			},
			120_000,
		);
	}

	it.skipIf(fixtureMissing(mediaFixture))(
		'moves an existing image, persists/reloads, and exports with a small binary-free snapshot',
		async () => {
			const source = await fixtureBytes(mediaFixture);
			expect(source.byteLength).toBeGreaterThan(5 * 1024 * 1024);
			const codec = new PptxCodec();
			const doc = new Doc({ gc: false });
			await codec.hydrate(doc, source);
			const original = await loadPptx(source);
			const slides = slideMaps(doc);
			for (const slide of slides) {
				slide.set('isDirty', false);
			}
			const index = slides.findIndex((slide) =>
				elementMaps(slide).some((element) => element.get('imagePath')),
			);
			expect(index).toBeGreaterThanOrEqual(0);
			const slide = slides[index];
			const image = elementMaps(slide).find((element) => element.get('imagePath'))!;
			const id = image.get('id');
			const x = Number(image.get('x'));
			const stateVector = encodeStateVector(doc);
			doc.transact(() => {
				image.set('x', x + 24);
				slide.set('isDirty', true);
			});
			expect(slide.get('isDirty')).toBeTruthy();
			expect(encodeStateAsUpdate(doc, stateVector).length).toBeLessThan(2048);
			// gc:false also exposes any source bytes inserted then deleted during seed.
			expect(encodeStateAsUpdate(doc).length).toBeLessThan(source.byteLength / 10);
			expect(JSON.stringify(doc.toJSON())).not.toContain('data:image/');
			expect(doc.getMap('pptx:meta').has('sourceBytes')).toBeFalsy();
			const restored = restoredDoc(doc);
			doc.destroy();
			// A restored document has no local source cache until explicitly registered.
			expect(restored.getMap('pptx:assets').size).toBe(0);
			const snapshot = encodeStateAsUpdate(restored);
			registerCollaborationSource(restored, original.slides);
			expect(encodeStateAsUpdate(restored)).toStrictEqual(snapshot);
			writeSlidesToYDoc(readSlidesFromYDoc(restored), restored, factories);
			expect(restored.getMap('pptx:assets').size).toBe(0);
			const output = await new PptxCodec().dehydrate(restored, source);
			const actual = await loadPptx(output);
			const moved = flattenElements(actual.slides[index].elements).find(
				(element) => element.id === id,
			)!;
			expect(moved.x).toBeCloseTo(x + 24, 3);
			const prior = flattenElements(original.slides[index].elements).find(
				(element) => element.id === id,
			)!;
			for (const key of [
				'imageData',
				'imagePath',
				'width',
				'height',
				'cropLeft',
				'cropTop',
				'cropRight',
				'cropBottom',
			]) {
				expect(Reflect.get(moved, key), `image field ${key}`).toStrictEqual(
					Reflect.get(prior, key),
				);
			}
			await expectPackagePreserved(source, output, [actual.slides[index].id]);
			// Logical deletion must not leave the native source in Yjs tombstones.
			restored.getArray('pptx:slides').delete(0, slides.length);
			expect(encodeStateAsUpdate(restored).length).toBeLessThan(source.length / 10);
			restored.destroy();
		},
		120_000,
	);
});

describe('canonical shared/tools schema', () => {
	it('rejects signed package export with SAVE_SIGNATURES_STRIPPED', async () => {
		const zip = await JSZip.loadAsync(await createTestPptxBytes(1));
		zip.file(
			'_xmlsignatures/origin.sigs',
			'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
		);
		zip.file('_xmlsignatures/sig1.xml', '<Signature xmlns="http://www.w3.org/2000/09/xmldsig#"/>');
		const relationships = await zip.file('_rels/.rels')!.async('string');
		zip.file(
			'_rels/.rels',
			relationships.replace(
				'</Relationships>',
				'<Relationship Id="rIdSig" Type="http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin" Target="_xmlsignatures/origin.sigs"/></Relationships>',
			),
		);
		const source = await zip.generateAsync({ type: 'uint8array' });
		const doc = new Doc();
		const codec = new PptxCodec();
		await codec.hydrate(doc, source);
		replaceText(slideMaps(doc)[0], 'Signed package edit');
		const snapshot = encodeStateAsUpdate(doc);
		await expect(codec.dehydrate(doc, source)).rejects.toThrow(/SAVE_SIGNATURES_STRIPPED/);
		expect(encodeStateAsUpdate(doc)).toStrictEqual(snapshot);
		doc.destroy();
	});

	it.skipIf(fixtureMissing('sample-deck.pptx'))(
		'table cell Y.Text survives shared, snapshot and tools export',
		async () => {
			const source = await fixtureBytes('sample-deck.pptx');
			const codec = new PptxCodec();
			const doc = new Doc();
			await codec.hydrate(doc, source);
			const slide = slideMaps(doc)[4];
			const element = elementMaps(slide).find((map) => map.get('type') === 'table')!;
			const table = element.get('tableData');
			expect(table instanceof YMap).toBeTruthy();
			const row = ((table as YMap<unknown>).get('rows') as YArray<YMap<unknown>>).get(0);
			const cell = (row.get('cells') as YArray<YMap<unknown>>).get(0);
			const text = cell.get('textBody') as Text;
			expect(text instanceof Text).toBeTruthy();
			text.delete(0, text.length);
			text.insert(0, 'Persisted table cell');
			slide.set('isDirty', true);
			const restored = restoredDoc(doc);
			const shared = readSlidesFromYDoc(restored)[4].elements.find(
				(item) => item.type === 'table',
			)!;
			expect(shared.tableData?.rows[0].cells[0].text).toBe('Persisted table cell');
			const output = await codec.dehydrate(restored, source);
			const actual = (await loadPptx(output)).slides[4].elements.find(
				(item) => item.type === 'table',
			)!;
			expect(actual.tableData?.rows[0].cells[0].text).toBe('Persisted table cell');
			await expectPackagePreserved(source, output, [slide.get('id') as string]);
			doc.destroy();
			restored.destroy();
		},
	);

	it('shared -> persisted snapshot -> tools export -> shared preserves text, styles and IDs', async () => {
		const source = await createTestPptxBytes(2);
		const codec = new PptxCodec();
		const doc = new Doc();
		await codec.hydrate(doc, source);
		const slides = readSlidesFromYDoc(doc);
		const element = slides[0].elements.find((item) => 'textSegments' in item)!;
		if ('textSegments' in element) {
			element.text = 'Shared schema edit';
			element.textSegments = [
				{ text: 'Shared schema edit', style: { bold: true, color: '#123456' } },
			];
		}
		slides[0].isDirty = true;
		writeSlidesToYDoc(slides, doc, factories);
		const reloaded = restoredDoc(doc);
		const output = await codec.dehydrate(reloaded, source);
		const exported = new Doc();
		await codec.hydrate(exported, output);
		const actual = readSlidesFromYDoc(exported);
		expect(actual.map((slide) => slide.id)).toStrictEqual(slides.map((slide) => slide.id));
		const edited = actual[0].elements.find((item) => item.id === element.id)!;
		expect(edited).toMatchObject({
			text: 'Shared schema edit',
			textSegments: [{ text: 'Shared schema edit', style: { bold: true, color: '#123456' } }],
		});
		expect(
			elementMaps(slideMaps(exported)[0])
				.find((map) => map.get('id') === element.id)!
				.get('textBody'),
		).toBeInstanceOf(Text);
		for (const item of [doc, reloaded, exported]) {
			item.destroy();
		}
	});

	it('rejects incompatible schema versions explicitly', async () => {
		const source = await createTestPptxBytes(1);
		const doc = new Doc();
		const codec = new PptxCodec();
		await codec.hydrate(doc, source);
		doc.getMap('pptx:meta').set('schemaVersion', -1);
		await expect(codec.dehydrate(doc, source)).rejects.toThrow(/schema|version/i);
		doc.destroy();
	});

	it('rejects export without its source and refuses unresolved source assets', async () => {
		const source = await createTestPptxBytes(1);
		const doc = new Doc();
		const codec = new PptxCodec();
		await codec.hydrate(doc, source);
		await expect(codec.dehydrate(doc, new Uint8Array())).rejects.toThrow(/source/i);
		const element = elementMaps(slideMaps(doc)[0])[0];
		element.set('_imgRef', `pptx-source:${'0'.repeat(64)}`);
		await expect(codec.dehydrate(doc, source)).rejects.toThrow(/source asset|unavailable/i);
		doc.destroy();
	});

	it('does not silently drop corrupted collaborative style JSON', async () => {
		const source = await createTestPptxBytes(1);
		const doc = new Doc();
		const codec = new PptxCodec();
		await codec.hydrate(doc, source);
		elementMaps(slideMaps(doc)[0])[0].set('_ts', '{invalid JSON');
		await expect(codec.dehydrate(doc, source)).rejects.toThrow(/JSON|style|corrupt|invalid/i);
		doc.destroy();
	});

	it('nested children are editable Y.Arrays and survive tools export', async () => {
		const source = await createTestPptxBytes(1);
		const codec = new PptxCodec();
		const doc = new Doc();
		await codec.hydrate(doc, source);
		const slides = readSlidesFromYDoc(doc);
		const children = slides[0].elements;
		slides[0].elements = [
			{ id: 'group-test', type: 'group', x: 0, y: 0, width: 960, height: 540, children },
		];
		slides[0].isDirty = true;
		writeSlidesToYDoc(slides, doc, factories);
		const group = (slideMaps(doc)[0].get('elements') as YArray<YMap<unknown>>).get(0);
		expect(group.get('children')).toBeInstanceOf(YArray);
		expect(group.has('_ch')).toBeFalsy();
		replaceText(slideMaps(doc)[0], 'Nested collaborative edit');
		const output = await codec.dehydrate(restoredDoc(doc), source);
		const loaded = await loadPptx(output);
		expect(loaded.slides[0].elements[0].type).toBe('group');
		expect(
			flattenElements(loaded.slides[0].elements).some(
				(item) => 'text' in item && item.text === 'Nested collaborative edit',
			),
		).toBeTruthy();
		doc.destroy();
	});
});
