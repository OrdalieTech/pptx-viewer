/**
 * collaboration-reconcile.ts: Granular Y.Doc reconciliation for collaborative
 * editing.
 *
 * `writeSlidesToYDoc` replaces the entire pptx:slides array on every write,
 * which makes concurrent edits collide at document granularity (last writer
 * wins for the whole deck). `reconcileSlidesInYDoc` instead diffs the desired
 * slide state against the live Y.Doc and only mutates what changed:
 *
 *  - slides and elements are matched by `id`; unchanged ones keep their Y.Map
 *    instance so concurrent field edits merge via Yjs
 *  - scalar / complex fields are compared and only set when different
 *  - textBody is edited in place (minimal char-level diff via
 *    collaboration-text-merge.ts) when its canonical decoded form differs;
 *    wholesale replacement is only a fallback
 *  - removed items are deleted; moves change order ranks while retaining maps
 *
 * All mutations run in a single transaction tagged with LOCAL_SYNC_ORIGIN (or
 * a caller-supplied origin) so observers can ignore their own writes.
 */

import type { PptxElement, PptxSlide } from 'pptx-viewer-core';

import { ASSET_ELEMENT_FIELDS, getAssetsMap, reconcileAssetFields } from './collaboration-assets';
import { reorderYArray } from './collaboration-order';
import { mapSourceAssets } from './collaboration-source';
import type { YArrayLike, YDocLike, YjsFactories, YMapLike } from './collaboration-sync';
import {
	COMPLEX_ELEMENT_FIELDS,
	COMPLEX_SLIDE_FIELDS,
	SCALAR_ELEMENT_KEYS,
	SCALAR_SLIDE_KEYS,
	writeElementToYMap,
	writeSlideToYMap,
	YDOC_SLIDES_KEY,
	assertCollaborationSchema,
	YDOC_META_KEY,
	YDOC_SCHEMA_VERSION,
} from './collaboration-sync';
import { reconcileTableData } from './collaboration-table';
import {
	decodeDelta,
	encodeSegmentsToDelta,
	encodeTextBody,
	isYTextLike,
} from './collaboration-text-codec';
import { isYTextEditable, mergeDeltaIntoYText } from './collaboration-text-merge';

/** Transaction origin used for local reconcile writes. */
export const LOCAL_SYNC_ORIGIN = 'pptx-viewer:local-sync';

function jsonEqual(a: unknown, b: unknown): boolean {
	return a === b || JSON.stringify(a) === JSON.stringify(b);
}

function reconcileScalars(
	ymap: YMapLike,
	rec: Record<string, unknown>,
	keys: ReadonlySet<string>,
): void {
	for (const key of keys) {
		if (key === 'backgroundImage' || ASSET_ELEMENT_FIELDS.has(key)) {
			continue;
		}
		const next = rec[key];
		const current = ymap.get(key);
		if (next === undefined) {
			if (current !== undefined) {
				ymap.delete(key);
			}
		} else if (!jsonEqual(current, next)) {
			ymap.set(key, next);
		}
	}
}

function reconcileComplexFields(
	ymap: YMapLike,
	rec: Record<string, unknown>,
	fields: Readonly<Record<string, string>>,
	assets: YMapLike,
): void {
	for (const [original, prefixed] of Object.entries(fields)) {
		const next =
			rec[original] === undefined
				? undefined
				: JSON.stringify(mapSourceAssets(rec[original], assets, true));
		const current = ymap.get(prefixed);
		if (next === undefined) {
			if (current !== undefined) {
				ymap.delete(prefixed);
			}
		} else if (current !== next) {
			ymap.set(prefixed, next);
		}
	}
}

/**
 * Bring `ymap.textBody` in line with `rec.textSegments`, preferring a
 * character-level in-place merge (so concurrent typing on one element
 * converges) and falling back to a wholesale Y.Text replacement.
 *
 * Exported for the live-patch channel (collaboration-live-patch.ts), which
 * writes interim editor text into the doc mid-edit and must use the exact same
 * merge path as this reconcile pass.
 */
export function reconcileElementTextBody(
	ymap: YMapLike,
	rec: Record<string, unknown>,
	factories: YjsFactories,
): void {
	const segments = rec.textSegments;
	const current = ymap.get('textBody');
	if (!Array.isArray(segments)) {
		if (current !== undefined) {
			ymap.delete('textBody');
		}
		return;
	}
	const desiredDelta = encodeSegmentsToDelta(segments);
	const desired = decodeDelta(desiredDelta);
	const existing = isYTextLike(current) ? decodeDelta(current.toDelta()) : undefined;
	if (existing !== undefined && jsonEqual(existing, desired)) {
		return;
	}
	// Prefer an in-place minimal edit so concurrent edits to the same text
	// element merge at character granularity instead of element-level LWW.
	if (isYTextEditable(current) && mergeDeltaIntoYText(current, desiredDelta)) {
		return;
	}
	const ytext = factories.createText();
	encodeTextBody(segments, ytext);
	ymap.set('textBody', ytext);
}

export function reconcileElementYMap(
	ymap: YMapLike,
	element: PptxElement,
	factories: YjsFactories,
	assets: YMapLike,
): void {
	const rec = element as unknown as Record<string, unknown>;
	reconcileScalars(ymap, rec, SCALAR_ELEMENT_KEYS);
	reconcileComplexFields(ymap, rec, COMPLEX_ELEMENT_FIELDS, assets);
	reconcileElementTextBody(ymap, rec, factories);
	reconcileAssetFields(rec.id as string, rec, ymap, assets);
	if (element.type === 'table' && element.tableData) {
		reconcileTableData(element.tableData, ymap, factories, element.id);
	}
	if (element.type === 'group') {
		let children = ymap.get('children');
		if (!isYArrayLike(children)) {
			children = factories.createArray();
			ymap.set('children', children);
		}
		reconcileYArrayById<PptxElement>(children as YArrayLike, element.children, {
			idOf: (child) => child.id,
			create: (child) => {
				const map = factories.createMap();
				writeElementToYMap(child, map, factories, assets);
				return map;
			},
			update: (map, child) => reconcileElementYMap(map, child, factories, assets),
		});
	}
}

interface ReconcileAdapter<T> {
	idOf: (item: T) => string | undefined;
	create: (item: T) => YMapLike;
	update: (ymap: YMapLike, item: T) => void;
}

function mapIdAt(arr: YArrayLike, index: number): string | undefined {
	const entry = arr.get(index) as YMapLike | undefined;
	if (!entry || typeof entry.get !== 'function') {
		return undefined;
	}
	const id = entry.get('id');
	return typeof id === 'string' ? id : undefined;
}

/**
 * Reconcile a Y.Array of Y.Maps against a desired item list, matching by id.
 * Every collaborative item requires a stable string ID.
 */
function reconcileYArrayById<T>(
	arr: YArrayLike,
	items: readonly T[],
	adapter: ReconcileAdapter<T>,
): void {
	const ids = items.map(adapter.idOf);
	if (ids.every((id): id is string => typeof id === 'string')) {
		if (new Set(ids).size !== ids.length) {
			throw new Error('Duplicate PPTX collaboration ID');
		}
		const desired = new Set(ids);
		for (let i = arr.length - 1; i >= 0; i--) {
			if (!desired.has(mapIdAt(arr, i) ?? '')) {
				arr.delete(i, 1);
			}
		}
		const existing = new Map(
			arr.toArray().map((entry) => {
				const map = entry as YMapLike;
				return [map.get('id'), map] as const;
			}),
		);
		items.forEach((item, index) => {
			const map = existing.get(ids[index]);
			if (map) {
				adapter.update(map, item);
			} else {
				const created = adapter.create(item);
				created.set('_order', arr.length);
				arr.push([created]);
			}
		});
		reorderYArray(arr, ids);
		return;
	}
	throw new Error('PPTX collaborative slides and elements require stable string IDs');
}

const isYArrayLike = (value: unknown): value is YArrayLike =>
	typeof value === 'object' &&
	value !== null &&
	typeof (value as YArrayLike).insert === 'function' &&
	typeof (value as YArrayLike).toArray === 'function';

export function reconcileSlideYMap(
	ymap: YMapLike,
	slide: PptxSlide,
	factories: YjsFactories,
	assets: YMapLike,
): void {
	const rec = slide as unknown as Record<string, unknown>;
	reconcileScalars(ymap, rec, SCALAR_SLIDE_KEYS);
	reconcileComplexFields(ymap, rec, COMPLEX_SLIDE_FIELDS, assets);
	reconcileAssetFields(slide.id, rec, ymap, assets);

	let elements = ymap.get('elements');
	if (!isYArrayLike(elements)) {
		elements = factories.createArray();
		ymap.set('elements', elements);
	}
	reconcileYArrayById<PptxElement>(elements as YArrayLike, slide.elements, {
		idOf: (el) => (typeof el.id === 'string' ? el.id : undefined),
		create: (el) => {
			const map = factories.createMap();
			writeElementToYMap(el, map, factories, assets);
			return map;
		},
		update: (map, el) => reconcileElementYMap(map, el, factories, assets),
	});
}

/**
 * Granular local -> Y.Doc sync: mutate only what changed, inside one
 * transaction tagged with `origin` (default LOCAL_SYNC_ORIGIN) so the
 * caller's own observer can skip the resulting events.
 */
export function reconcileSlidesInYDoc(
	slides: readonly PptxSlide[],
	ydoc: YDocLike,
	factories: YjsFactories,
	origin: unknown = LOCAL_SYNC_ORIGIN,
): void {
	const assets = getAssetsMap(ydoc);
	ydoc.transact(() => {
		assertCollaborationSchema(ydoc);
		if (ydoc.getMap(YDOC_META_KEY).get('schemaVersion') === undefined) {
			ydoc.getMap(YDOC_META_KEY).set('schemaVersion', YDOC_SCHEMA_VERSION);
		}
		const arr = ydoc.getArray(YDOC_SLIDES_KEY);
		reconcileYArrayById<PptxSlide>(arr, slides, {
			idOf: (slide) => (typeof slide.id === 'string' ? slide.id : undefined),
			create: (slide) => {
				const map = factories.createMap();
				writeSlideToYMap(slide, map, factories, assets);
				return map;
			},
			update: (map, slide) => reconcileSlideYMap(map, slide, factories, assets),
		});
	}, origin);
}
