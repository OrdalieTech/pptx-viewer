import { XMLBuilder, XMLParser } from 'fast-xml-parser';
import JSZip from 'jszip';
import { PptxHandler } from 'pptx-viewer-core';
import type { PptxElement, PptxSlide, XmlObject } from 'pptx-viewer-core';

const URI = 'urn:pptx-viewer:collaboration:identities:1';
const NODE = 'cv:identities';
type TableIds = { id: string; cells: string[] }[];
type Identity = { slideId: string; elements: Map<string, string>; tables: Map<string, TableIds> };
const array = (value: unknown): XmlObject[] =>
	value === undefined ? [] : ((Array.isArray(value) ? value : [value]) as XmlObject[]);
const template = (element: PptxElement): boolean => /^(layout-|master-)/.test(element.id);

function elements(slide: PptxSlide): PptxElement[] {
	const result: PptxElement[] = [];
	const visit = (items: PptxElement[]): void => {
		for (const element of items) {
			if (template(element)) continue;
			result.push(element);
			if (element.type === 'group') visit(element.children);
		}
	};
	visit(slide.elements);
	return result;
}

function identity(slide: PptxSlide): Identity | undefined {
	const root = slide.rawXml?.['p:sld'] as XmlObject | undefined;
	const extList = root?.['p:extLst'] as XmlObject | undefined;
	const matches = array(extList?.['p:ext']).filter((ext) => ext['@_uri'] === URI);
	if (matches.length > 1) throw new Error('Duplicate PPTX collaboration identity extension');
	if (matches.length === 0) return undefined;
	const metadata = matches[0]?.[NODE] as XmlObject | undefined;
	if (
		!metadata ||
		metadata['@_version'] !== '1' ||
		typeof metadata['@_slideId'] !== 'string' ||
		!metadata['@_slideId']
	) {
		throw new Error('Unsupported PPTX native collaboration identity version');
	}
	const mapped = new Map<string, string>();
	const ids = new Set<string>();
	for (const entry of array(metadata['cv:element'])) {
		const shapeId = entry['@_shapeId'],
			id = entry['@_id'];
		if (
			typeof shapeId !== 'string' ||
			typeof id !== 'string' ||
			!id ||
			mapped.has(shapeId) ||
			ids.has(id)
		) {
			throw new Error('Invalid PPTX native collaboration element identity');
		}
		mapped.set(shapeId, id);
		ids.add(id);
	}
	const tables = new Map<string, TableIds>();
	for (const entry of array(metadata['cv:table'])) {
		const shapeId = entry['@_shapeId'];
		const rows: unknown = JSON.parse(String(entry['@_rows']));
		const used = new Set<string>();
		const validId = (id: unknown): boolean => {
			if (typeof id !== 'string' || !id || used.has(id)) return false;
			used.add(id);
			return true;
		};
		if (
			typeof shapeId !== 'string' ||
			tables.has(shapeId) ||
			!Array.isArray(rows) ||
			!rows.every(
				(row) => row && validId(row.id) && Array.isArray(row.cells) && row.cells.every(validId),
			)
		)
			throw new Error('Invalid PPTX native table identities');
		tables.set(shapeId, rows);
	}
	return { slideId: metadata['@_slideId'], elements: mapped, tables };
}

/** Restore stable IDs on parsed models before constructing their collaborative Yjs state. */
export function restorePptxIdentities(slides: PptxSlide[]): void {
	const ids = new Set<string>();
	for (const slide of slides) {
		const mapping = identity(slide);
		if (mapping) {
			slide.id = mapping.slideId;
			for (const element of elements(slide)) {
				const restored =
					element.shapeId === undefined ? undefined : mapping.elements.get(String(element.shapeId));
				if (restored) element.id = restored;
				const rows = mapping.tables.get(String(element.shapeId));
				if (rows) {
					if (element.type !== 'table') {
						throw new Error('PPTX native table identity targets a non-table element');
					}
					const data = element.tableData;
					if (
						!data ||
						rows.length !== data.rows.length ||
						rows.some((row, i) => row.cells.length !== data.rows[i].cells.length)
					)
						throw new Error('PPTX native table identity dimensions differ');
					rows.forEach((row, i) => {
						data.rows[i].collaborationId = row.id;
						row.cells.forEach((id, j) => {
							data.rows[i].cells[j].collaborationId = id;
						});
					});
				}
			}
		}
		if (ids.has(slide.id)) throw new Error('Duplicate PPTX collaborative slide ID');
		ids.add(slide.id);
	}
}

/**
 * Resolve native source paths before save, then stamp only IDs that differ from
 * the parser's native identity. The original Y.Doc is never mutated.
 */
export function preparePptxIdentities(
	slides: PptxSlide[],
	nativeSourceSlides: PptxSlide[],
): { finish: (bytes: Uint8Array) => Promise<Uint8Array> } {
	const sources = new Map<string, PptxSlide>();
	for (const source of nativeSourceSlides) {
		const id = identity(source)?.slideId ?? source.id;
		if (sources.has(id)) throw new Error('Duplicate native PPTX collaborative slide ID');
		sources.set(id, source);
	}
	const plans = slides.map((slide) => {
		const slideId = slide.id;
		const source = sources.get(slideId);
		if (source) {
			slide.id = source.id;
			slide.rId = source.rId;
		}
		if (slide.sourceSlideId)
			slide.sourceSlideId = sources.get(slide.sourceSlideId)?.id ?? slide.sourceSlideId;
		const mapping = source ? identity(source) : undefined;
		const nativeById = new Map(
			(source ? elements(source) : []).map((element) => [
				mapping?.elements.get(String(element.shapeId)) ?? element.id,
				element,
			]),
		);
		const all = elements(slide);
		const used = new Set<string>(['1']); // p:spTree's root group owns ID 1.
		for (const element of all) if (element.shapeId !== undefined) used.add(String(element.shapeId));
		for (const element of nativeById.values())
			if (element.shapeId !== undefined) used.add(String(element.shapeId));
		let next = 2;
		const allocated = new Set<string>();
		const stableIds = new Set<string>();
		const records = all.map((element) => {
			if (stableIds.has(element.id)) throw new Error('Duplicate PPTX collaborative element ID');
			stableIds.add(element.id);
			const prior = nativeById.get(element.id);
			if (prior?.shapeId !== undefined) element.shapeId = prior.shapeId;
			if (element.shapeId === undefined || allocated.has(String(element.shapeId))) {
				while (used.has(String(next))) next++;
				element.shapeId = String(next++);
				used.add(element.shapeId);
			}
			allocated.add(String(element.shapeId));
			return { element, id: element.id };
		});
		return { slide, slideId, records };
	});

	return {
		async finish(bytes) {
			const loaded = await new PptxHandler().load(bytes.slice().buffer as ArrayBuffer);
			const parsedByPath = new Map(loaded.slides.map((slide) => [slide.id, slide]));
			let zip: JSZip | undefined;
			const parser = new XMLParser({
				ignoreAttributes: false,
				parseTagValue: false,
				trimValues: false,
				preserveOrder: true,
			});
			const builder = new XMLBuilder({
				ignoreAttributes: false,
				format: false,
				preserveOrder: true,
			});
			for (const plan of plans) {
				const parsed = parsedByPath.get(plan.slide.id);
				if (!parsed) throw new Error(`Export lost slide ${plan.slideId}`);
				const nativeByShape = new Map(
					elements(parsed).map((element) => [String(element.shapeId), element]),
				);
				const entries: XmlObject[] = [];
				const tables = new Map<string, TableIds>();
				for (const record of plan.records) {
					const shapeId = String(record.element.shapeId);
					const native = nativeByShape.get(shapeId);
					if (!native) throw new Error(`Export lost element ${record.id} (shape ${shapeId})`);
					if (native.id !== record.id) entries.push({ '@_shapeId': shapeId, '@_id': record.id });
					if (record.element.type === 'table' && record.element.tableData) {
						const rows = record.element.tableData.rows.map((row, i) => ({
							id: row.collaborationId ?? `${record.id}:row:${i}`,
							cells: row.cells.map(
								(cell, j) => cell.collaborationId ?? `${record.id}:cell:${i}:${j}`,
							),
						}));
						if (
							rows.some(
								(row, i) =>
									row.id !== `${record.id}:row:${i}` ||
									row.cells.some((id, j) => id !== `${record.id}:cell:${i}:${j}`),
							)
						)
							tables.set(shapeId, rows);
					}
				}
				const previous = identity(parsed);
				if (!entries.length && !tables.size && parsed.id === plan.slideId && !previous) continue;
				if (
					previous?.slideId === plan.slideId &&
					previous.elements.size === entries.length &&
					JSON.stringify([...previous.tables]) === JSON.stringify([...tables]) &&
					entries.every(
						(entry) => previous.elements.get(String(entry['@_shapeId'])) === entry['@_id'],
					)
				)
					continue;
				zip ??= await JSZip.loadAsync(bytes);
				const file = zip.file(plan.slide.id);
				if (!file) throw new Error(`Export lost native slide part ${plan.slide.id}`);
				// Preserve child order, including interleaved drawing types and text runs.
				const xml = parser.parse(await file.async('string')) as XmlObject[];
				const root = xml.find((node) => 'p:sld' in node)?.['p:sld'] as XmlObject[];
				if (!Array.isArray(root)) throw new Error('Exported slide has no p:sld root');
				let extList = root.find((node) => 'p:extLst' in node);
				if (!extList) {
					extList = { 'p:extLst': [] };
					root.push(extList);
				}
				const extensions = array(extList['p:extLst']).filter(
					(ext) => (ext[':@'] as XmlObject)?.['@_uri'] !== URI,
				);
				extensions.push({
					':@': { '@_uri': URI },
					'p:ext': [
						{
							[NODE]: [
								...entries.map((attributes) => ({ 'cv:element': [], ':@': attributes })),
								...[...tables].map(([shapeId, rows]) => ({
									'cv:table': [],
									':@': { '@_shapeId': shapeId, '@_rows': JSON.stringify(rows) },
								})),
							],
							':@': { '@_xmlns:cv': URI, '@_version': '1', '@_slideId': plan.slideId },
						},
					],
				});
				extList['p:extLst'] = extensions;
				zip.file(plan.slide.id, builder.build(xml));
			}
			return zip ? zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }) : bytes;
		},
	};
}
