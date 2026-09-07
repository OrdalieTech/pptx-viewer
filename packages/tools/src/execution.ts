import { PptxHandler } from 'pptx-viewer-core';
import type { PptxData } from 'pptx-viewer-core';
import {
	readSlidesFromYDoc,
	registerCollaborationSource,
	reconcileSlidesInYDoc,
} from 'pptx-viewer-shared/collaboration';
import type { YDocLike } from 'pptx-viewer-shared/collaboration';
import { Doc, applyUpdate, encodeStateAsUpdate, encodeStateVector } from 'yjs';

import { PptxCodec, yjsFactories } from './codec/pptx-codec.js';
import type { ExecutionContext, ToolContext, ToolResult } from './types.js';

const PPTX_MIME_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
const loadedRooms = new WeakMap<
	Uint8Array,
	{ doc: Doc; state: Uint8Array; sourceSlides: PptxData['slides'] }
>();

/**
 * Load PptxData from a file, checking collaboration rooms first.
 * An active room is read from an isolated snapshot, including unsaved edits.
 */
export async function loadPresentation(
	filePath: string,
	execCtx: ExecutionContext,
): Promise<{ pptxData: PptxData; rawBytes: Uint8Array }> {
	const rawBytes = (await execCtx.filesystem.readFile(filePath)).slice();

	// Capture before parsing yields so later human edits remain independent.
	if (execCtx.collaboration) {
		const room = execCtx.collaboration.getRoom(filePath);
		if (room) {
			const codec = execCtx.collaboration.getCodec(filePath);
			if (!codec) {
				throw new Error('Collaboration codec is unavailable');
			}
			if (codec) {
				const state = encodeStateAsUpdate(room.ydoc);
				const snapshot = new Doc();
				try {
					applyUpdate(snapshot, state);
					new PptxCodec().assertCompatibleYDoc(snapshot);
					const pptxData = await new PptxHandler().load(rawBytes.slice().buffer as ArrayBuffer);
					const sourceSlides = pptxData.slides;
					registerCollaborationSource(snapshot as unknown as YDocLike, sourceSlides);
					pptxData.slides = readSlidesFromYDoc(snapshot as unknown as YDocLike);
					loadedRooms.set(rawBytes, { doc: room.ydoc, state, sourceSlides });
					return { pptxData, rawBytes };
				} finally {
					snapshot.destroy();
				}
			}
		}
	}

	// Fallback: parse from raw bytes on disk
	const handler = new PptxHandler();
	const pptxData = await handler.load(rawBytes.slice().buffer as ArrayBuffer);
	return { pptxData, rawBytes };
}

/**
 * Save mutated PptxData back: routes through collaboration Y.Doc or disk.
 */
export async function savePresentation(
	filePath: string,
	pptxData: PptxData,
	rawBytes: Uint8Array,
	execCtx: ExecutionContext,
): Promise<{ savedToDisk: boolean; routedThroughCollaboration: boolean }> {
	// Reconcile against the tool's original snapshot, then merge only its delta.
	if (execCtx.collaboration) {
		const room = execCtx.collaboration.getRoom(filePath);
		if (room) {
			const codec = execCtx.collaboration.getCodec(filePath);
			if (!codec) {
				throw new Error('Collaboration codec is unavailable');
			}
			if (codec) {
				const loaded = loadedRooms.get(rawBytes);
				if (!loaded || loaded.doc !== room.ydoc) {
					throw new Error('Collaboration room changed; reload the presentation before saving');
				}
				const staged = new Doc();
				try {
					applyUpdate(staged, loaded.state);
					const vector = encodeStateVector(staged);
					registerCollaborationSource(staged as unknown as YDocLike, loaded.sourceSlides);
					const origin = execCtx.collaboration.agentOrigin(execCtx.agentName ?? 'pptx-tool');
					reconcileSlidesInYDoc(
						pptxData.slides,
						staged as unknown as YDocLike,
						yjsFactories,
						origin,
					);
					applyUpdate(room.ydoc, encodeStateAsUpdate(staged, vector), origin);
					loadedRooms.delete(rawBytes);
				} finally {
					staged.destroy();
				}
				return { savedToDisk: false, routedThroughCollaboration: true };
			}
		}
	}
	if (loadedRooms.has(rawBytes)) {
		throw new Error('Collaboration room is unavailable; refusing a local binary save');
	}
	const handler = new PptxHandler();
	await handler.load(rawBytes.slice().buffer as ArrayBuffer);
	const outputBytes = await handler.save(pptxData.slides, {
		headerFooter: pptxData.headerFooter,
	});

	// Non-collaboration: write to disk
	await execCtx.filesystem.writeFile(filePath, outputBytes);

	// Update viewer if available
	if (execCtx.viewer) {
		execCtx.viewer.replaceContent(filePath, outputBytes, {
			markDirty: false,
			mimeType: PPTX_MIME_TYPE,
		});
	}

	return { savedToDisk: true, routedThroughCollaboration: false };
}

/**
 * Execute a pure tool function with full collaboration-aware load/save pipeline.
 * Handles snapshot → tool → incremental room delta routing automatically.
 */
export async function executeToolWithContext<T>(
	filePath: string,
	execCtx: ExecutionContext,
	toolFn: (ctx: ToolContext) => ToolResult<T> | Promise<ToolResult<T>>,
): Promise<T & { savedToDisk?: boolean; routedThroughCollaboration?: boolean }> {
	const { pptxData, rawBytes } = await loadPresentation(filePath, execCtx);

	const toolResult = await toolFn({ pptxData });

	if (toolResult.dirty) {
		const saveResult = await savePresentation(filePath, toolResult.pptxData, rawBytes, execCtx);
		return { ...toolResult.result, ...saveResult };
	}

	return toolResult.result as T & {
		savedToDisk?: boolean;
		routedThroughCollaboration?: boolean;
	};
}
