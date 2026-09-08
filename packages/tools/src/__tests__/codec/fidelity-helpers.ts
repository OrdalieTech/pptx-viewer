import { createHash } from 'node:crypto';
import { existsSync } from 'node:fs';
import { readFile } from 'node:fs/promises';
import { posix } from 'node:path';

import { XMLParser, XMLValidator } from 'fast-xml-parser';
import JSZip from 'jszip';
import { PptxHandler } from 'pptx-viewer-core';
import type { PptxElement } from 'pptx-viewer-core';
import { expect } from 'vitest';
import { Array as YArray, Doc, Map as YMap, Text, applyUpdate, encodeStateAsUpdate } from 'yjs';

export const factories = {
	createMap: () => new YMap(),
	createArray: () => new YArray(),
	createText: () => new Text(),
};
export const fixtureUrl = (name: string) =>
	new URL(`../../../../../e2e/fixtures/${name}`, import.meta.url);
export const fixtureMissing = (name: string) => !existsSync(fixtureUrl(name));
export const fixtureBytes = async (name: string) =>
	new Uint8Array(await readFile(fixtureUrl(name)));
export const loadPptx = (bytes: Uint8Array) => new PptxHandler().load(bytes.slice().buffer);
export const xmlParser = new XMLParser({
	ignoreAttributes: false,
	parseTagValue: false,
	parseAttributeValue: false,
	ignoreDeclaration: true,
	trimValues: false,
	// Indentation between package nodes is not slide text.
	tagValueProcessor: (tag: string, value: string) => (tag === 'a:t' ? value : value.trim()),
});

export function restoredDoc(doc: Doc): Doc {
	const restored = new Doc();
	applyUpdate(restored, encodeStateAsUpdate(doc));
	return restored;
}

export function slideMaps(doc: Doc): YMap<unknown>[] {
	return doc.getArray<YMap<unknown>>('pptx:slides').toArray();
}

export function elementMaps(slide: YMap<unknown>): YMap<unknown>[] {
	const elements = slide.get('elements') as YArray<YMap<unknown>>;
	const visit = (maps: YMap<unknown>[]): YMap<unknown>[] =>
		maps.flatMap((map) => {
			const children = map.get('children');
			return [map, ...(children instanceof YArray ? visit(children.toArray()) : [])];
		});
	return visit(elements.toArray());
}

export function flattenElements(elements: PptxElement[]): PptxElement[] {
	return elements.flatMap((element) => [
		element,
		...(element.type === 'group' ? flattenElements(element.children) : []),
	]);
}

export function replaceText(slide: YMap<unknown>, replacement: string): string {
	const candidates = elementMaps(slide).filter((map) => map.get('textBody') instanceof Text);
	const element =
		candidates.find((map) => !/^(layout-|master-)/.test(String(map.get('id')))) ?? candidates[0];
	expect(element, 'fixture must contain editable text').toBeDefined();
	const body = element!.get('textBody') as Text;
	const attributes = body.toDelta()[0]?.attributes;
	body.delete(0, body.length);
	body.insert(0, replacement, attributes);
	slide.set('isDirty', true);
	expect(slide.get('isDirty')).toBeTruthy();
	return element!.get('id') as string;
}

export async function expectPackagePreserved(
	source: Uint8Array,
	output: Uint8Array,
	editedSlides: string[],
): Promise<void> {
	const before = await JSZip.loadAsync(source);
	const after = await JSZip.loadAsync(output);
	const sourceHasComments = before.file(/^ppt\/comments\/comment\d+\.xml$/).length > 0;
	for (const [path, part] of Object.entries(before.files)) {
		if (part.dir) {
			continue;
		}
		if (!sourceHasComments && path === 'ppt/commentAuthors.xml' && !after.file(path)) {
			continue;
		}
		expect(after.file(path), `missing package part ${path}`).not.toBeNull();
		if (editedSlides.includes(path)) {
			continue;
		}
		const original = await part.async('uint8array');
		const actual = await after.file(path)!.async('uint8array');
		if (path.endsWith('.xml') || path.endsWith('.rels')) {
			const beforeXml = xmlParser.parse(new TextDecoder().decode(original));
			const afterXml = xmlParser.parse(new TextDecoder().decode(actual));
			if (path === 'docProps/core.xml') {
				// Save bookkeeping may update revision/time and supply a missing author.
				const oldProperties = beforeXml['cp:coreProperties'];
				const properties = afterXml['cp:coreProperties'];
				expect(Number(properties['cp:revision'])).toBe(
					Number(oldProperties['cp:revision'] ?? 0) + 1,
				);
				expect(Number.isFinite(Date.parse(properties['dcterms:modified']['#text']))).toBeTruthy();
				if (oldProperties['cp:lastModifiedBy'] === undefined) {
					expect(properties['cp:lastModifiedBy']).toBe('pptx');
					delete properties['cp:lastModifiedBy'];
				}
				if (
					oldProperties['cp:lastModifiedBy'] === '' &&
					properties['cp:lastModifiedBy'] === 'pptx'
				) {
					delete oldProperties['cp:lastModifiedBy'];
					delete properties['cp:lastModifiedBy'];
				}
				for (const field of ['cp:revision', 'dcterms:modified']) {
					delete oldProperties[field];
					delete properties[field];
				}
			}
			if (path === 'docProps/app.xml') {
				// Some authored decks omit derived counts; save fills them in.
				const slides = (await loadPptx(source)).slides;
				const counts = {
					Slides: slides.length,
					HiddenSlides: slides.filter((slide) => slide.hidden).length,
					Notes: slides.filter((slide) => slide.notes?.trim()).length,
				};
				for (const [field, count] of Object.entries(counts)) {
					const root = Object.keys(beforeXml).find(
						(key) => key === 'Properties' || key.endsWith(':Properties'),
					)!;
					const prefix = root.includes(':') ? `${root.split(':')[0]}:` : '';
					const key = `${prefix}${field}`;
					if (beforeXml[root][key] === undefined) {
						expect(afterXml[root][key], `derived ${field}`).toBe(String(count));
						delete afterXml[root][key];
					}
				}
			}
			if (!sourceHasComments && path === '[Content_Types].xml') {
				const overrides = beforeXml.Types.Override;
				beforeXml.Types.Override = (Array.isArray(overrides) ? overrides : [overrides]).filter(
					(entry) => entry?.['@_PartName'] !== '/ppt/commentAuthors.xml',
				);
			}
			if (!sourceHasComments && path === 'ppt/_rels/presentation.xml.rels') {
				for (const xml of [beforeXml, afterXml]) {
					const relationships = xml.Relationships.Relationship;
					xml.Relationships.Relationship = (
						Array.isArray(relationships) ? relationships : [relationships]
					)
						.filter(
							(entry) =>
								sourceHasComments || !String(entry?.['@_Type']).endsWith('/commentAuthors'),
						)
						.map((entry) => ({
							...entry,
							'@_Target': String(entry['@_Target']).startsWith('/')
								? String(entry['@_Target']).slice(1)
								: posix.normalize(posix.join('ppt', String(entry['@_Target']))),
						}))
						.sort((a, b) => String(a['@_Id']).localeCompare(String(b['@_Id'])));
				}
			}
			expect(afterXml, `unedited XML ${path}`).toStrictEqual(beforeXml);
		} else {
			expect(createHash('sha256').update(actual).digest('hex'), `binary part ${path}`).toBe(
				createHash('sha256').update(original).digest('hex'),
			);
		}
	}
	await expectValidRelationships(after);
}

export async function expectValidRelationships(zip: JSZip): Promise<void> {
	for (const [path, file] of Object.entries(zip.files)) {
		if (file.dir || (!path.endsWith('.xml') && !path.endsWith('.rels'))) {
			continue;
		}
		const xml = await file.async('string');
		expect(XMLValidator.validate(xml), `invalid XML ${path}`).toBeTruthy();
		if (!path.endsWith('.rels')) {
			continue;
		}
		const relationships = xmlParser.parse(xml).Relationships?.Relationship ?? [];
		const ids = new Set<string>();
		for (const relation of Array.isArray(relationships) ? relationships : [relationships]) {
			expect(
				ids.has(relation['@_Id']),
				`duplicate relationship ${path}:${relation['@_Id']}`,
			).toBeFalsy();
			ids.add(relation['@_Id']);
			if (relation['@_TargetMode'] === 'External') {
				continue;
			}
			const base = path === '_rels/.rels' ? '' : posix.dirname(posix.dirname(path));
			const target = decodeURIComponent(relation['@_Target'].split('#')[0]);
			const resolved = target.startsWith('/')
				? target.slice(1)
				: posix.normalize(posix.join(base, target));
			expect(zip.file(resolved), `dangling relationship ${path} -> ${resolved}`).not.toBeNull();
		}
	}
}
