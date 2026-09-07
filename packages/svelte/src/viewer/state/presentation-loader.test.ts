import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import type { PptxHandler, PptxSlide } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import { resolveLazyImages } from './loader-helpers';
import { PresentationLoader } from './presentation-loader.svelte';

/**
 * Load-pipeline wiring tests against a real `.pptx` fixture: the full
 * `PptxHandler.load` path (ZIP -> XML -> theme/master/layout resolution ->
 * chart enrichment) plus this package's media/image URL resolution and
 * reactive state commit.
 */

// Vitest runs with cwd = packages/svelte; the fixture lives at the repo root.
const FIXTURE = resolve(process.cwd(), '../../e2e/fixtures/sample-deck.pptx');

function readFixture(): Uint8Array {
	return new Uint8Array(readFileSync(FIXTURE));
}

describe('presentationLoader', () => {
	it('indexes lazy picture URLs for collaborative slides that only retain imagePath', async () => {
		const imagePath = 'ppt/media/image1.png';
		const imageData = 'data:image/png;base64,AAAA';
		const slides = [
			{
				id: 'ppt/slides/slide1.xml',
				elements: [{ id: 'picture-1', type: 'picture', imagePath }],
			},
		] as unknown as PptxSlide[];
		const resolvedUrls = new Map<string, string>();
		const handler = {
			getImageData: async () => imageData,
		} as unknown as PptxHandler;

		const resolved = await resolveLazyImages(handler, slides, resolvedUrls);

		expect(resolved[0].elements[0]).toMatchObject({ imagePath, imageData });
		expect(resolvedUrls.get(imagePath)).toBe(imageData);
	});

	it('loads a real presentation into reactive state', async () => {
		const loader = new PresentationLoader();
		expect(loader.loadCount).toBe(0);

		await loader.load(readFixture());

		expect(loader.loading).toBeFalsy();
		expect(loader.error).toBeNull();
		expect(loader.isEncrypted).toBeFalsy();
		expect(loader.handler).not.toBeNull();
		expect(loader.loadCount).toBe(1);
		expect(loader.slides.length).toBeGreaterThan(0);
		expect(loader.canvasSize.width).toBeGreaterThan(0);
		expect(loader.canvasSize.height).toBeGreaterThan(0);
		expect(Array.isArray(loader.embeddedFonts)).toBeTruthy();
		expect(loader.digitalSignatureCount).toBeGreaterThanOrEqual(0);
		expect(loader.hasDigitalSignatures).toBe(loader.digitalSignatureCount > 0);
		// Every slide carries an id + elements array (the viewer's contract).
		for (const slide of loader.slides) {
			expect(slide.id).toBeTruthy();
			expect(Array.isArray(slide.elements)).toBeTruthy();
		}

		loader.dispose();
		expect(loader.handler).toBeNull();
	});

	it('reports a parse error for garbage bytes', async () => {
		const loader = new PresentationLoader();
		await loader.load(new Uint8Array([1, 2, 3, 4, 5]));
		expect(loader.loading).toBeFalsy();
		expect(loader.error).toBeTruthy();
		expect(loader.loadCount).toBe(0);
		loader.dispose();
	});

	it('ignores stale loads when a newer one is issued', async () => {
		const loader = new PresentationLoader();
		const first = loader.load(readFixture());
		const second = loader.load(readFixture());
		await Promise.all([first, second]);
		// Only the latest load commits; the stale one is discarded.
		expect(loader.loadCount).toBe(1);
		expect(loader.slides.length).toBeGreaterThan(0);
		loader.dispose();
	});
});
