import { readFile } from 'node:fs/promises';

import JSZip from 'jszip';
import { PptxHandler } from 'pptx-viewer-core';
import type { PptxSlide } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import { preparePptxIdentities, restorePptxIdentities } from './pptx-identities';

const fixture = () =>
	readFile(new URL('../../../../e2e/fixtures/master-views.pptx', import.meta.url));

describe('native collaborative identity', () => {
	it('keeps new slides and elements stable over two exports and reloads', async () => {
		const input = await fixture();
		const handler = new PptxHandler();
		const source = await handler.load(
			input.buffer.slice(input.byteOffset, input.byteOffset + input.byteLength),
		);
		const slides: PptxSlide[] = [
			...structuredClone(source.slides),
			{
				id: 'collaborative-slide',
				elements: [
					{
						id: 'collaborative-table',
						type: 'table',
						x: 10,
						y: 100,
						width: 200,
						height: 50,
						tableData: {
							columnWidths: [200],
							rows: [
								{
									collaborationId: 'inserted-row',
									cells: [{ collaborationId: 'stable-cell', text: 'Cell', style: { bold: true } }],
								},
							],
						},
					},
					{
						id: 'collaborative-text',
						type: 'text',
						x: 10,
						y: 10,
						width: 200,
						height: 50,
						text: 'Stable element',
						textSegments: [{ text: 'Stable element', style: { bold: true } }],
					},
				],
			},
		];
		for (const slide of slides) slide.isDirty = true;
		const stableState = structuredClone(slides);
		const firstPlan = preparePptxIdentities(slides, source.slides);
		const first = await firstPlan.finish(await handler.save(slides));
		const secondHandler = new PptxHandler();
		const parsed = await secondHandler.load(first.slice().buffer);
		const firstNativePaths = parsed.slides.map((slide) => slide.id);
		const restored = structuredClone(parsed.slides);
		restorePptxIdentities(restored);
		expect(restored.at(-1)?.id).toBe('collaborative-slide');
		expect(
			restored.at(-1)?.elements.find((element) => element.type === 'table')?.tableData?.rows[0],
		).toMatchObject({
			collaborationId: 'inserted-row',
			cells: [{ collaborationId: 'stable-cell', text: 'Cell' }],
		});
		expect(restored.at(-1)?.elements.some((element) => element.id === 'collaborative-text')).toBe(
			true,
		);
		const secondPlan = preparePptxIdentities(stableState, parsed.slides);
		const second = await secondPlan.finish(await secondHandler.save(stableState));
		const again = await new PptxHandler().load(second.slice().buffer);
		expect(again.slides.map((slide) => slide.id)).toEqual(firstNativePaths);
		restorePptxIdentities(again.slides);
		expect(
			again.slides.at(-1)?.elements.find((element) => element.type === 'table')?.tableData?.rows[0]
				.cells[0].collaborationId,
		).toBe('stable-cell');
		expect(
			again.slides.at(-1)?.elements.find((element) => element.id === 'collaborative-text')?.text,
		).toBe('Stable element');
	}, 30_000);

	it('does not touch unchanged native slide parts', async () => {
		const input = await fixture();
		const handler = new PptxHandler();
		const source = await handler.load(
			input.buffer.slice(input.byteOffset, input.byteOffset + input.byteLength),
		);
		const slides = structuredClone(source.slides);
		for (const slide of slides) slide.isDirty = false;
		const plan = preparePptxIdentities(slides, source.slides);
		const saved = await handler.save(slides);
		const stamped = await plan.finish(saved);
		const before = await JSZip.loadAsync(saved),
			after = await JSZip.loadAsync(stamped);
		for (const slide of slides)
			expect(await after.file(slide.id)!.async('string')).toBe(
				await before.file(slide.id)!.async('string'),
			);
	}, 30_000);
});
