import { digest } from 'lib0/hash/sha256';
import type { PptxSlide } from 'pptx-viewer-core';

import type { YDocLike, YMapLike } from './collaboration-sync';

// Source payloads are local render/export inputs, never shared Yjs state.
const sources = new WeakMap<YMapLike, Map<string, string>>();
const references = new WeakMap<YMapLike, Map<string, string>>();
export const SOURCE_ASSET_PREFIX = 'pptx-source:';
const binaryFields = new Set([
	'imageData',
	'svgData',
	'mediaData',
	'posterFrameData',
	'oleEmbeddedData',
	'previewImageData',
	'modelData',
	'previewImage',
	'posterImage',
	'backgroundImage',
]);

export function registerCollaborationSource(doc: YDocLike, slides: readonly PptxSlide[]): void {
	const assets = doc.getMap('pptx:assets');
	const values = new Map<string, string>();
	const refs = new Map<string, string>();
	const visit = (
		value: unknown,
		field = '',
		parent: Record<string, unknown> = {},
		slidePath = '',
	): void => {
		if (typeof value === 'string' && binaryFields.has(field) && value.length > 0) {
			if (refs.has(value)) {
				return;
			}
			let identity = value;
			const pathField: Record<string, string> = {
				imageData: 'imagePath',
				svgData: 'svgPath',
				previewImageData: 'previewImage',
				posterFrameData: 'posterFramePath',
				mediaData: 'mediaPath',
				modelData: 'modelPath',
				oleEmbeddedData: 'oleTarget',
			};
			const path = parent[pathField[field]];
			if (typeof path === 'string' && path && !/^(?:blob:|data:|https?:)/u.test(path)) {
				const resolved = path.startsWith('ppt/')
					? path
					: new URL(path, `https://pptx.invalid/${slidePath}`).pathname.slice(1);
				identity = `package:${resolved}`;
			} else if (value.startsWith('blob:')) {
				throw new Error(`PPTX source asset ${field} has no stable package path`);
			}
			const hash = Array.from(digest(new TextEncoder().encode(identity)), (b) =>
				b.toString(16).padStart(2, '0'),
			).join('');
			const ref = SOURCE_ASSET_PREFIX + hash;
			values.set(ref, value);
			refs.set(value, ref);
		} else if (Array.isArray(value)) {
			for (const item of value) {
				visit(item, '', {}, slidePath);
			}
		} else if (value && typeof value === 'object') {
			const record = value as Record<string, unknown>;
			const currentSlide =
				Array.isArray(record.elements) && typeof record.id === 'string' ? record.id : slidePath;
			for (const [key, child] of Object.entries(value)) {
				visit(child, key, record, currentSlide);
			}
		}
	};
	visit(slides);
	sources.set(assets, values);
	references.set(assets, refs);
}

export function sourceAssetReference(assets: YMapLike, value: string): string | undefined {
	return value.startsWith(SOURCE_ASSET_PREFIX) ? value : references.get(assets)?.get(value);
}

export function resolveSourceAsset(assets: YMapLike, value: string): string {
	return sources.get(assets)?.get(value) ?? value;
}

export function mapSourceAssets(value: unknown, assets: YMapLike, encode: boolean): unknown {
	if (typeof value === 'string') {
		return encode
			? (sourceAssetReference(assets, value) ?? value)
			: resolveSourceAsset(assets, value);
	}
	if (Array.isArray(value)) {
		return value.map((child) => mapSourceAssets(child, assets, encode));
	}
	if (value && typeof value === 'object') {
		return Object.fromEntries(
			Object.entries(value).map(([key, child]) => [key, mapSourceAssets(child, assets, encode)]),
		);
	}
	return value;
}

export function assertSourceAssetsResolved(value: unknown): void {
	if (typeof value === 'string' && value.startsWith(SOURCE_ASSET_PREFIX)) {
		throw new Error(`PPTX source asset is unavailable: ${value}`);
	}
	if (Array.isArray(value)) {
		for (const child of value) {
			assertSourceAssetsResolved(child);
		}
	} else if (value && typeof value === 'object') {
		for (const child of Object.values(value)) {
			assertSourceAssetsResolved(child);
		}
	}
}
