import type { PptxSlide } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';
import { Doc } from 'yjs';

import {
	mapSourceAssets,
	registerCollaborationSource,
	sourceAssetReference,
	resolveSourceAsset,
} from './collaboration-source';

describe('source background identity', () => {
	it.each([false, true])(
		'preserves exported background identities after a slide path changes (array: %s)',
		(array) => {
			const doc = new Doc();
			const ref = `pptx-source:${'a'.repeat(64)}`;
			const extension = {
				'@_uri': 'urn:pptx-viewer:collaboration:identities:1',
				'cv:identities': { '@_backgroundRef': ref },
			};
			try {
				registerCollaborationSource(doc, [
					{
						id: 'ppt/slides/slide2.xml',
						elements: [],
						backgroundImage: 'blob:reopened',
						rawXml: { 'p:sld': { 'p:extLst': { 'p:ext': array ? [extension] : extension } } },
					} as PptxSlide,
				]);
				expect(resolveSourceAsset(doc.getMap('pptx:assets'), ref)).toBe('blob:reopened');
				expect(sourceAssetReference(doc.getMap('pptx:assets'), 'blob:reopened')).toBe(ref);
				expect(doc.getMap('pptx:assets').size).toBe(0);
			} finally {
				doc.destroy();
			}
		},
	);
	it('resolves an exported background against a fresh load with different blob URLs', () => {
		const first = new Doc();
		const next = new Doc();
		try {
			const slide = (backgroundImage: string) =>
				({ id: 'ppt/slides/slide1.xml', elements: [], backgroundImage }) as PptxSlide;
			registerCollaborationSource(first, [slide('blob:first')]);
			registerCollaborationSource(next, [slide('blob:next')]);
			const ref = sourceAssetReference(first.getMap('pptx:assets'), 'blob:first')!;
			expect(ref).toBe(sourceAssetReference(next.getMap('pptx:assets'), 'blob:next'));
			expect(resolveSourceAsset(next.getMap('pptx:assets'), ref)).toBe('blob:next');
		} finally {
			first.destroy();
			next.destroy();
		}
	});
});

describe('mapSourceAssets', () => {
	it('decodes parsed JSON as it is without a source, and resolves references with one', () => {
		const doc = new Doc();
		try {
			const assets = doc.getMap('pptx:assets');
			const ref = `pptx-source:${'b'.repeat(64)}`;
			const json = JSON.stringify({ fill: { image: ref }, list: [ref, 'plain', 3, null], 2: true });
			expect(mapSourceAssets(JSON.parse(json), assets, false)).toStrictEqual(JSON.parse(json));
			expect(JSON.stringify(mapSourceAssets(JSON.parse(json), assets, false))).toBe(json);

			registerCollaborationSource(doc, [
				{ id: 'ppt/slides/slide1.xml', elements: [], backgroundImage: 'blob:image' } as PptxSlide,
			]);
			const known = sourceAssetReference(assets, 'blob:image')!;
			expect(
				mapSourceAssets({ fill: { image: known }, list: [known, ref] }, assets, false),
			).toStrictEqual({ fill: { image: 'blob:image' }, list: ['blob:image', ref] });
			expect(mapSourceAssets({ list: ['blob:image', 'plain'] }, assets, true)).toStrictEqual({
				list: [known, 'plain'],
			});
		} finally {
			doc.destroy();
		}
	});
});
