/**
 * collaboration-sync.ts: Framework-agnostic CRDT sync utilities for the
 * pptx-viewer collaboration stack (Yjs backend).
 *
 * Exports:
 *  - Structural Yjs interfaces (no hard yjs import - bindings pass live instances)
 *  - YjsFactories: factory interface bindings implement using `new Y.Map()` etc.
 *  - writeElementToYMap / readElementFromYMap: PptxElement <-> YMapLike
 *  - writeSlideToYMap / readSlideFromYMap: PptxSlide <-> YMapLike
 *  - writeSlidesToYDoc / readSlidesFromYDoc: PptxSlide[] <-> Y.Doc
 *  - observeYDocSlides: register a change listener on the pptx:slides array
 *  - re-exports of the text codec (collaboration-text-codec.ts)
 *
 * Y.Doc schema:
 *   pptx:slides  - Y.Array of slide Y.Maps
 *   Each slide Y.Map has scalar keys + `_`-prefixed JSON blobs + `elements`
 *   Each element Y.Map has scalar keys + `_`-prefixed JSON blobs + `textBody`
 *   textBody is a Y.Text with one delta-op per TextSegment
 *
 * This is also the canonical tools/codec schema. Source assets are local inputs.
 *
 * Prefer `reconcileSlidesInYDoc` (collaboration-reconcile.ts) over
 * `writeSlidesToYDoc` for live editing: it updates only what changed instead
 * of replacing the whole slides array, so concurrent edits merge per
 * slide/element/field rather than colliding at document granularity.
 */

import type { PptxSlide, PptxElement } from 'pptx-viewer-core';

import {
	ASSET_ELEMENT_FIELDS,
	getAssetsMap,
	isAssetRefKey,
	isAssetVersionKey,
	readAssetFields,
	writeAssetFields,
} from './collaboration-assets';
import { orderedYMaps } from './collaboration-order';
import {
	assertCollaborationSchema,
	YDOC_META_KEY,
	YDOC_SCHEMA_VERSION,
	YDOC_SLIDES_KEY,
	SCALAR_ELEMENT_KEYS,
	COMPLEX_ELEMENT_FIELDS,
	SCALAR_SLIDE_KEYS,
	COMPLEX_SLIDE_FIELDS,
} from './collaboration-schema';
import type {
	YMapLike,
	YArrayLike,
	YDocLike,
	YjsFactories,
	YDeepObserver,
} from './collaboration-schema';
import { mapSourceAssets } from './collaboration-source';
import { readTableData, writeTableData } from './collaboration-table';
import { encodeTextBody, decodeTextBody, isYTextLike } from './collaboration-text-codec';

export * from './collaboration-assets';
export * from './collaboration-text-codec';
export * from './collaboration-source';
export * from './collaboration-order';
export * from './collaboration-table';
export * from './collaboration-schema';

const REV_COMPLEX_ELEMENT: Record<string, string> = Object.fromEntries(
	Object.entries(COMPLEX_ELEMENT_FIELDS).map(([k, v]) => [v, k]),
);

const REV_COMPLEX_SLIDE: Record<string, string> = Object.fromEntries(
	Object.entries(COMPLEX_SLIDE_FIELDS).map(([k, v]) => [v, k]),
);

// ---------------------------------------------------------------------------
// Element serialization
// ---------------------------------------------------------------------------

export function writeElementToYMap(
	element: PptxElement,
	ymap: YMapLike,
	factories: YjsFactories,
	assets: YMapLike,
): void {
	const rec = element as unknown as Record<string, unknown>;
	for (const [key, value] of Object.entries(rec)) {
		if (value === undefined || ASSET_ELEMENT_FIELDS.has(key)) {
			continue;
		}
		if (SCALAR_ELEMENT_KEYS.has(key)) {
			ymap.set(key, value);
		} else if (key === 'tableData') {
			writeTableData(
				value as import('pptx-viewer-core').PptxTableData,
				ymap,
				factories,
				element.id,
			);
		} else if (key === 'children' && Array.isArray(value)) {
			const children = factories.createArray();
			for (const [index, child] of value.entries()) {
				const childMap = factories.createMap();
				writeElementToYMap(child, childMap, factories, assets);
				childMap.set('_order', index);
				children.push([childMap]);
			}
			ymap.set('children', children);
		} else if (key === 'textSegments') {
			if (Array.isArray(value)) {
				const ytext = factories.createText();
				encodeTextBody(value, ytext);
				ymap.set('textBody', ytext);
			}
		} else if (COMPLEX_ELEMENT_FIELDS[key]) {
			ymap.set(COMPLEX_ELEMENT_FIELDS[key], JSON.stringify(mapSourceAssets(value, assets, true)));
		}
	}
	writeAssetFields(rec.id as string, rec, ymap, assets);
}

export function readElementFromYMap(ymap: YMapLike, assets: YMapLike): PptxElement {
	const element: Record<string, unknown> = {};
	ymap.forEach((value: unknown, key: string) => {
		if (key === '_order') {
			return;
		}
		if (key === 'tableData') {
			element.tableData = readTableData(ymap);
		} else if (key === 'children') {
			element.children = orderedYMaps(value as YArrayLike).map((child) =>
				readElementFromYMap(child, assets),
			);
		} else if (key === 'textBody') {
			if (isYTextLike(value)) {
				element.textSegments = decodeTextBody(value);
			}
		} else if (isAssetRefKey(key) || isAssetVersionKey(key)) {
			// Ref pointers are resolved by readAssetFields below; version
			// counters are an internal sync token, never a PptxElement field.
		} else if (REV_COMPLEX_ELEMENT[key]) {
			element[REV_COMPLEX_ELEMENT[key]] = mapSourceAssets(
				JSON.parse(value as string),
				assets,
				false,
			);
		} else {
			element[key] = value;
		}
	});
	readAssetFields(ymap, assets, element);
	if (Array.isArray(element.textSegments)) {
		element.text = element.textSegments.map((segment: { text: string }) => segment.text).join('');
	}
	return element as unknown as PptxElement;
}

// ---------------------------------------------------------------------------
// Slide serialization
// ---------------------------------------------------------------------------

export function writeSlideToYMap(
	slide: PptxSlide,
	ymap: YMapLike,
	factories: YjsFactories,
	assets: YMapLike,
): void {
	const rec = slide as unknown as Record<string, unknown>;
	for (const key of SCALAR_SLIDE_KEYS) {
		if (key === 'backgroundImage' || ASSET_ELEMENT_FIELDS.has(key)) {
			continue;
		}
		if (rec[key] !== undefined) {
			ymap.set(key, rec[key]);
		}
	}
	for (const [original, prefixed] of Object.entries(COMPLEX_SLIDE_FIELDS)) {
		if (rec[original] !== undefined) {
			ymap.set(prefixed, JSON.stringify(mapSourceAssets(rec[original], assets, true)));
		}
	}
	const elemArr = factories.createArray();
	for (const [index, el] of slide.elements.entries()) {
		const elemMap = factories.createMap();
		writeElementToYMap(el, elemMap, factories, assets);
		elemMap.set('_order', index);
		elemArr.push([elemMap]);
	}
	ymap.set('elements', elemArr);
	writeAssetFields(rec.id as string, rec, ymap, assets);
}

export function readSlideFromYMap(ymap: YMapLike, assets: YMapLike): PptxSlide {
	const slide: Record<string, unknown> = {};
	for (const key of SCALAR_SLIDE_KEYS) {
		const v = ymap.get(key);
		if (v !== undefined) {
			slide[key] = v;
		}
	}
	for (const [prefixed, original] of Object.entries(REV_COMPLEX_SLIDE)) {
		const v = ymap.get(prefixed) as string | undefined;
		if (v !== undefined) {
			slide[original] = mapSourceAssets(JSON.parse(v), assets, false);
		}
	}
	const elemArr = ymap.get('elements') as YArrayLike | undefined;
	const elements: PptxElement[] = [];
	if (elemArr) {
		for (const element of orderedYMaps(elemArr)) {
			elements.push(readElementFromYMap(element, assets));
		}
	}
	slide.elements = elements;
	readAssetFields(ymap, assets, slide);
	return slide as unknown as PptxSlide;
}

// ---------------------------------------------------------------------------
// Y.Doc-level helpers
// ---------------------------------------------------------------------------

/**
 * Replace the full slides array in the Y.Doc. Coarse: prefer
 * `reconcileSlidesInYDoc` for live editing; this remains suitable for
 * one-shot seeding of an empty document.
 */
export function writeSlidesToYDoc(
	slides: PptxSlide[],
	ydoc: YDocLike,
	factories: YjsFactories,
	origin?: unknown,
): void {
	const assets = getAssetsMap(ydoc);
	ydoc.transact(() => {
		assertCollaborationSchema(ydoc);
		ydoc.getMap(YDOC_META_KEY).set('schemaVersion', YDOC_SCHEMA_VERSION);
		const arr = ydoc.getArray(YDOC_SLIDES_KEY);
		if (arr.length > 0) {
			arr.delete(0, arr.length);
		}
		for (const [index, slide] of slides.entries()) {
			const ymap = factories.createMap();
			writeSlideToYMap(slide, ymap, factories, assets);
			ymap.set('_order', index);
			arr.push([ymap]);
		}
	}, origin);
}

export function readSlidesFromYDoc(ydoc: YDocLike): PptxSlide[] {
	assertCollaborationSchema(ydoc);
	const assets = getAssetsMap(ydoc);
	const arr = ydoc.getArray(YDOC_SLIDES_KEY);
	const slides: PptxSlide[] = [];
	for (const slide of orderedYMaps(arr)) {
		slides.push(readSlideFromYMap(slide, assets));
	}
	return slides;
}

/**
 * Observe (deeply) the pptx:slides array. The handler receives the Yjs
 * events plus the transaction, so callers can skip their own writes by
 * checking `transaction.origin` (see LOCAL_SYNC_ORIGIN in
 * collaboration-reconcile.ts).
 */
export function observeYDocSlides(ydoc: YDocLike, onChange: YDeepObserver): () => void {
	const arr = ydoc.getArray(YDOC_SLIDES_KEY);
	arr.observeDeep(onChange);
	return () => arr.unobserveDeep(onChange);
}
