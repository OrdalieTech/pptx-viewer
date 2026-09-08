import { PptxHandler } from 'pptx-viewer-core';
import {
	assertCollaborationSchema,
	assertSourceAssetsResolved,
	readSlidesFromYDoc,
	registerCollaborationSource,
	writeSlidesToYDoc,
	YDOC_SCHEMA_VERSION,
} from 'pptx-viewer-shared/collaboration';
import type { YDocLike, YjsFactories } from 'pptx-viewer-shared/collaboration';
import { Doc as YDoc, Array as YArray, Map as YMap, Text as YText } from 'yjs';

import { preparePptxIdentities, restorePptxIdentities } from './pptx-identities';

export {
	SCALAR_ELEMENT_KEYS,
	COMPLEX_ELEMENT_FIELDS as COMPLEX_FIELD_MAP,
	SCALAR_SLIDE_KEYS,
	COMPLEX_SLIDE_FIELDS as COMPLEX_SLIDE_FIELD_MAP,
} from 'pptx-viewer-shared/collaboration';

export const ORIGIN_FILE_LOAD = 'file-load';
export const yjsFactories: YjsFactories = {
	createMap: () => new YMap(),
	createArray: () => new YArray(),
	createText: () => new YText(),
};

function assertEmptySeedTarget(ydoc: YDoc): void {
	if (ydoc.getArray('pptx:slides').length || ydoc.getMap('pptx:meta').size) {
		throw new Error(
			'PPTX seed requires an empty document; existing collaboration must not be replaced',
		);
	}
}

/** The native package is an explicit export input, never a hidden Yjs payload. */
export interface FormatCodec {
	readonly formatId: string;
	readonly extensions: string[];
	hydrate: (ydoc: YDoc, bytes: Uint8Array, origin?: string) => Promise<void>;
	dehydrate: (ydoc: YDoc, baseSourcePptx: Uint8Array) => Promise<Uint8Array>;
	observe: (ydoc: YDoc, onChange: () => void) => () => void;
}

export class PptxCodec implements FormatCodec {
	readonly formatId = 'pptx';
	readonly extensions = ['.pptx'];

	assertCompatibleYDoc(ydoc: YDoc): void {
		assertCollaborationSchema(ydoc as unknown as YDocLike);
		if (ydoc.getMap('pptx:meta').get('schemaVersion') !== YDOC_SCHEMA_VERSION) {
			throw new Error('Unsupported PPTX collaboration schema: missing schemaVersion');
		}
	}

	async hydrate(ydoc: YDoc, bytes: Uint8Array, origin = ORIGIN_FILE_LOAD): Promise<void> {
		assertEmptySeedTarget(ydoc);
		if (!bytes.byteLength) {
			throw new Error('PPTX source is empty');
		}
		const handler = new PptxHandler();
		const data = await handler.load(bytes.slice().buffer as ArrayBuffer);
		// Parsing yields: a remote snapshot or another seed may have arrived meanwhile.
		assertEmptySeedTarget(ydoc);
		const doc = ydoc as unknown as YDocLike;
		// Resolve relative asset paths while slide IDs are still native package paths.
		registerCollaborationSource(doc, data.slides);
		restorePptxIdentities(data.slides);
		ydoc.transact(() => {
			const meta = ydoc.getMap('pptx:meta');
			meta.set('schemaVersion', YDOC_SCHEMA_VERSION);
			for (const field of ['width', 'height', 'widthEmu', 'heightEmu'] as const) {
				if (data[field] !== undefined) {
					meta.set(field, data[field]);
				}
			}
			writeSlidesToYDoc(data.slides, doc, yjsFactories, origin);
		}, origin);
	}

	async dehydrate(ydoc: YDoc, baseSourcePptx: Uint8Array): Promise<Uint8Array> {
		this.assertCompatibleYDoc(ydoc);
		if (!baseSourcePptx?.byteLength) {
			throw new Error('PPTX source is required for export');
		}
		const handler = new PptxHandler();
		const source = await handler.load(baseSourcePptx.slice().buffer as ArrayBuffer);
		const doc = ydoc as unknown as YDocLike;
		registerCollaborationSource(doc, source.slides);
		const slides = readSlidesFromYDoc(doc);
		if (!slides.length) {
			throw new Error('Cannot export a presentation without slides');
		}
		assertSourceAssetsResolved(slides);
		// Yjs geometry/text mutations do not depend on a viewer's local dirty flag.
		// Compare canonical projections to preserve untouched slide XML.
		const baseline = new YDoc();
		try {
			const baselineSlides = structuredClone(source.slides);
			restorePptxIdentities(baselineSlides);
			registerCollaborationSource(baseline as unknown as YDocLike, source.slides);
			writeSlidesToYDoc(baselineSlides, baseline as unknown as YDocLike, yjsFactories);
			const original = new Map(
				readSlidesFromYDoc(baseline as unknown as YDocLike).map((slide) => [slide.id, slide]),
			);
			for (const slide of slides) {
				const prior = original.get(slide.id);
				slide.isDirty =
					JSON.stringify({ ...slide, isDirty: undefined }) !==
					JSON.stringify({ ...prior, isDirty: undefined });
			}
			const identities = preparePptxIdentities(slides, source.slides);
			const output = await handler.save(slides);
			const warnings = handler
				.getCompatibilityWarnings()
				.filter((warning) => warning.scope === 'save');
			if (warnings.length) {
				throw new Error(
					`PPTX export would lose fidelity: ${warnings.map((warning) => `${warning.code}: ${warning.message}`).join('; ')}`,
				);
			}
			return await identities.finish(output);
		} finally {
			baseline.destroy();
		}
	}

	observe(ydoc: YDoc, onChange: () => void): () => void {
		const roots = [
			ydoc.getArray('pptx:slides'),
			ydoc.getMap('pptx:meta'),
			ydoc.getMap('pptx:assets'),
		];
		for (const root of roots) {
			root.observeDeep(onChange);
		}
		return () => {
			for (const root of roots) {
				root.unobserveDeep(onChange);
			}
		};
	}
}
