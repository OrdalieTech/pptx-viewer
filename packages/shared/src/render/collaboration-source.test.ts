import type { PptxSlide } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';
import { Doc } from 'yjs';

import {
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
