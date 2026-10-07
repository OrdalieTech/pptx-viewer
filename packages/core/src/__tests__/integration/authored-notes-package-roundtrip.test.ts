import { XMLParser } from 'fast-xml-parser';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { PresentationBuilder } from '../../core/builders/sdk/PresentationBuilder';
import { PptxHandler } from '../../core/PptxHandler';

describe('authored notes package round-trip', () => {
	it('creates notes master and slide OPC references for a new presentation', async () => {
		const { handler, data, createSlide } = await PresentationBuilder.create();
		const slide = createSlide('Blank').build();
		slide.notes = 'A newly authored speaker note';
		data.slides.push(slide);

		const saved = await handler.save(data.slides);
		const zip = await JSZip.loadAsync(saved);
		const requiredParts = [
			'ppt/notesMasters/notesMaster1.xml',
			'ppt/notesMasters/_rels/notesMaster1.xml.rels',
			'ppt/notesSlides/notesSlide1.xml',
			'ppt/notesSlides/_rels/notesSlide1.xml.rels',
		];
		for (const path of requiredParts) {
			expect(zip.file(path), `${path} should exist`).not.toBeNull();
		}

		const presentation = await zip.file('ppt/presentation.xml')!.async('string');
		expect(presentation).toContain('p:notesMasterIdLst');
		const presentationRels = await zip.file('ppt/_rels/presentation.xml.rels')!.async('string');
		expect(presentationRels).toContain('relationships/notesMaster');
		expect(presentationRels).toContain('notesMasters/notesMaster1.xml');

		const slideRels = await zip.file('ppt/slides/_rels/slide1.xml.rels')!.async('string');
		expect(slideRels).toContain('relationships/notesSlide');
		const notesRels = await zip.file('ppt/notesSlides/_rels/notesSlide1.xml.rels')!.async('string');
		expect(notesRels).toContain('relationships/notesMaster');
		expect(notesRels).toContain('relationships/slide');

		const contentTypes = await zip.file('[Content_Types].xml')!.async('string');
		expect(contentTypes).toContain('/ppt/notesMasters/notesMaster1.xml');
		expect(contentTypes).toContain('/ppt/notesSlides/notesSlide1.xml');

		const masterRelsPath = 'ppt/notesMasters/_rels/notesMaster1.xml.rels';
		const parser = new XMLParser({ ignoreAttributes: false });
		const privateTheme = async (archive: JSZip) => {
			const rels = parser.parse(await archive.file(masterRelsPath)!.async('string'));
			const target = rels.Relationships.Relationship['@_Target'];
			const path = new URL(
				target,
				'http://package/ppt/notesMasters/notesMaster1.xml',
			).pathname.slice(1);
			expect(path).not.toBe('ppt/theme/theme1.xml');
			await expect(archive.file(path)!.async('uint8array')).resolves.toStrictEqual(
				await archive.file('ppt/theme/theme1.xml')!.async('uint8array'),
			);
			return path;
		};
		await privateTheme(zip);

		// Existing stored decks still have the old shared relationship.
		zip.file(
			masterRelsPath,
			(await zip.file(masterRelsPath)!.async('string')).replace(
				/Target="[^"]*"/,
				'Target="../theme/theme1.xml"',
			),
		);
		const legacy = await zip.generateAsync({ type: 'arraybuffer' });
		const reloader = new PptxHandler();
		const reloaded = await reloader.load(legacy);
		expect(reloaded.slides[0].notes).toBe(slide.notes);
		expect(reloaded.notesMaster?.path).toBe('ppt/notesMasters/notesMaster1.xml');

		const resaved = await reloader.save(reloaded.slides, { notesMaster: reloaded.notesMaster });
		const restored = await JSZip.loadAsync(resaved);
		const theme = await privateTheme(restored);
		const finalLoader = new PptxHandler();
		const final = await finalLoader.load(resaved.buffer as ArrayBuffer);
		expect(final.slides[0].notes).toBe(slide.notes);
		expect(final.notesMaster).toBeDefined();
		const again = await finalLoader.save(final.slides, { notesMaster: final.notesMaster });
		await expect(privateTheme(await JSZip.loadAsync(again))).resolves.toBe(theme);
	});

	it.each(['case collision', 'case reference', 'new handout'])(
		'keeps the notes theme private on the first export: %s',
		async (scenario) => {
			const { handler, data, createSlide } = await PresentationBuilder.create();
			const slide = createSlide('Blank').build();
			slide.notes = 'Preserved notes';
			data.slides.push(slide);
			const zip = await JSZip.loadAsync(await handler.save(data.slides));
			const relsPath = 'ppt/notesMasters/_rels/notesMaster1.xml.rels';
			if (scenario !== 'new handout') {
				const target =
					scenario === 'case reference' ? '../THEME/THEME1.xml' : '../theme/theme1.xml';
				zip.file(
					relsPath,
					(await zip.file(relsPath)!.async('string')).replace(
						/Target="[^"]*"/,
						`Target="${target}"`,
					),
				);
			}
			if (scenario === 'case collision') {
				zip.file(
					'ppt/theme/Theme2.xml',
					await zip.file('ppt/theme/theme2.xml')!.async('uint8array'),
				);
				zip.remove('ppt/theme/theme2.xml');
				zip.file(
					'[Content_Types].xml',
					(await zip.file('[Content_Types].xml')!.async('string')).replaceAll(
						'/ppt/theme/theme2.xml',
						'/ppt/theme/Theme2.xml',
					),
				);
			}
			// The handout writer chooses the first theme in package order.
			const ordered = new JSZip();
			if (scenario === 'new handout') {
				ordered.file(
					'ppt/theme/theme2.xml',
					await zip.file('ppt/theme/theme2.xml')!.async('uint8array'),
				);
			}
			for (const [path, file] of Object.entries(zip.files)) {
				if (!file.dir) {
					ordered.file(path, await file.async('uint8array'));
				}
			}
			const reloader = new PptxHandler();
			const loaded = await reloader.load(await ordered.generateAsync({ type: 'arraybuffer' }));
			const options =
				scenario === 'new handout'
					? { handoutMaster: { path: 'ppt/handoutMasters/handoutMaster1.xml' } }
					: undefined;
			const exported = await JSZip.loadAsync(await reloader.save(loaded.slides, options));
			const names = Object.keys(exported.files)
				.filter((path) => !exported.files[path].dir)
				.map((path) => path.toLowerCase());
			expect(new Set(names).size).toBe(names.length);
			const parser = new XMLParser({ ignoreAttributes: false });
			const themeTargets = async (path: string) => {
				const rels = parser.parse(await exported.file(path)!.async('string'));
				return [rels.Relationships.Relationship]
					.flat()
					.filter((rel) => rel['@_Type'].endsWith('/theme'))
					.map((rel) => rel['@_Target'].toLowerCase());
			};
			const [notesTheme] = await themeTargets(relsPath);
			for (const path of Object.keys(exported.files).filter((candidate) =>
				/^ppt\/(slideMasters|handoutMasters)\/_rels\/.*\.rels$/.test(candidate),
			)) {
				await expect(themeTargets(path)).resolves.not.toContain(notesTheme);
			}
			expect(loaded.slides[0].notes).toBe(slide.notes);
		},
	);
});
