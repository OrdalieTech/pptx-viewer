/**
 * useYjsDocumentSync -- Syncs PptxSlide[] state with a Yjs Y.Doc using the
 * granular `pptx:slides` Y.Array structure (one Y.Map per slide, per-element
 * Y.Maps, and Y.Text for textSegments). This matches the schema defined by
 * PptxCodec in packages/tools so all bindings and the codec are interoperable.
 *
 * Write-back (Area 3): when the collaboration role is `'owner'`, the hook
 * debounces Y.Doc changes and calls `config.onWriteBack` with the serialized
 * PPTX bytes so the host can persist a durable snapshot.
 */

import { PptxHandler } from 'pptx-viewer-core';
import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import type { CollaborationConfig, YjsFactories, YTransactionLike } from 'pptx-viewer-shared';
import {
	reconcileSlidesInYDoc,
	LOCAL_SYNC_ORIGIN,
	readSlidesFromYDoc,
	registerCollaborationSource,
	failCollaborationSession,
	observeYDocSlides,
} from 'pptx-viewer-shared';
import { useCallback, useEffect, useRef } from 'react';
import type { Doc as YDoc } from 'yjs';

import { buildSaveSlides } from '../../utils/template-editing';

const WRITE_BACK_DEBOUNCE_DEFAULT_MS = 5_000;

export interface UseYjsDocumentSyncInput {
	/** The Yjs document (from useYjsProvider). null when not collaborating. */
	doc: YDoc | null;
	/** Current slides state. */
	slides: PptxSlide[];
	/** Separated master/layout (template) elements, merged back on write-back. */
	templateElementsBySlideId: Record<string, PptxElement[]>;
	/** React state setter for slides. */
	setSlides: React.Dispatch<React.SetStateAction<PptxSlide[]>>;
	/** Whether collaboration is active (status === 'connected'). */
	isConnected: boolean;
	/**
	 * Whether the provider completed its initial document sync. Local -> doc
	 * writes are gated on this so a late joiner never seeds its bootstrap deck
	 * into a room whose real content has not arrived yet. Defaults to true for
	 * callers that manage sync readiness themselves.
	 */
	isSynced?: boolean;
	/** Collaboration config (for role and write-back). */
	config?: Pick<CollaborationConfig, 'role' | 'onWriteBack' | 'writeBackDebounceMs' | 'onstatus'>;
	/**
	 * Return the source PPTX bytes for write-back serialization. Only called
	 * when role === 'owner' and onWriteBack is set.
	 */
	getSourceBytes?: () => Uint8Array | null;
	/**
	 * Monotonic counter bumped each time the content-load pipeline finishes
	 * applying a parsed deck to viewer state. A local load that lands while the
	 * shared doc already holds slides (a late joiner's bootstrap deck parsing
	 * after the room state arrived) would silently clobber the synced slides;
	 * each bump re-adopts the doc's slides when the room has content.
	 */
	loadVersion?: number;
}

export function useYjsDocumentSync({
	doc,
	slides,
	templateElementsBySlideId,
	setSlides,
	isConnected,
	isSynced = true,
	config,
	getSourceBytes,
	loadVersion = 0,
}: UseYjsDocumentSyncInput): void {
	const isApplyingRemoteRef = useRef(false);
	const lastSyncedRef = useRef('');
	const hasInitializedRef = useRef(false);
	const writeBackTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
	const factoriesRef = useRef<YjsFactories | null>(null);
	const lastLoadVersionRef = useRef(loadVersion);
	const failedRef = useRef(false);
	useEffect(() => {
		failedRef.current = false;
	}, [doc]);
	const readRemoteSlides = useCallback(
		(input: Parameters<typeof readSlidesFromYDoc>[0]) => {
			try {
				return readSlidesFromYDoc(input);
			} catch (cause) {
				failedRef.current = true;
				const error = cause instanceof Error ? cause : new Error(String(cause));
				failCollaborationSession(input, error);
				config?.onstatus?.('error', error);
				return [];
			}
		},
		[config],
	);
	const sourceSlidesRef = useRef(slides);
	if (lastLoadVersionRef.current !== loadVersion) {
		sourceSlidesRef.current = slides;
	}
	useEffect(() => {
		if (doc) {
			registerCollaborationSource(
				doc as unknown as Parameters<typeof registerCollaborationSource>[0],
				sourceSlidesRef.current,
			);
		}
	}, [doc, loadVersion]);

	const getFactories = useCallback(async (): Promise<YjsFactories> => {
		if (factoriesRef.current) {
			return factoriesRef.current;
		}
		const Y = await import('yjs');
		const factories: YjsFactories = {
			createMap: () => new Y.Map(),
			createArray: () => new Y.Array(),
			createText: () => new Y.Text(),
		};
		factoriesRef.current = factories;
		return factories;
	}, []);

	const scheduleWriteBack = useCallback(() => {
		if (!config?.onWriteBack || config.role !== 'owner' || !doc) {
			return;
		}
		if (writeBackTimerRef.current !== null) {
			clearTimeout(writeBackTimerRef.current);
		}
		const debounceMs = config.writeBackDebounceMs ?? WRITE_BACK_DEBOUNCE_DEFAULT_MS;
		writeBackTimerRef.current = setTimeout(async () => {
			writeBackTimerRef.current = null;
			if (!doc || !config.onWriteBack) {
				return;
			}
			const sourceBytes = getSourceBytes?.();
			if (!sourceBytes) {
				return;
			}
			try {
				const handler = new PptxHandler();
				await handler.load(sourceBytes.buffer as ArrayBuffer);
				const currentSlides = readRemoteSlides(
					doc as unknown as Parameters<typeof readSlidesFromYDoc>[0],
				);
				if (failedRef.current) {
					return;
				}
				const slidesToSave = buildSaveSlides(currentSlides, templateElementsBySlideId);
				const bytes = await handler.save(slidesToSave);
				config.onWriteBack(bytes);
			} catch {
				/* write-back failures are non-fatal */
			}
		}, debounceMs);
	}, [doc, config, getSourceBytes, templateElementsBySlideId, readRemoteSlides]);

	useEffect(() => {
		if (loadVersion === lastLoadVersionRef.current) {
			return;
		}
		lastLoadVersionRef.current = loadVersion;
		if (!doc || !isConnected) {
			return;
		}
		const docSlides = readRemoteSlides(doc as unknown as Parameters<typeof readSlidesFromYDoc>[0]);
		if (docSlides.length === 0) {
			return;
		}
		lastSyncedRef.current = JSON.stringify(docSlides);
		isApplyingRemoteRef.current = true;
		setSlides(docSlides);
		isApplyingRemoteRef.current = false;
	}, [loadVersion, doc, isConnected, setSlides, readRemoteSlides]);

	useEffect(() => {
		if (
			!isConnected ||
			!isSynced ||
			!doc ||
			failedRef.current ||
			config?.role === 'viewer' ||
			isApplyingRemoteRef.current ||
			slides.length === 0
		) {
			return;
		}

		if (!hasInitializedRef.current) {
			const remote = readRemoteSlides(doc as unknown as Parameters<typeof readSlidesFromYDoc>[0]);
			if (failedRef.current) {
				return;
			}
			if (remote.length) {
				lastSyncedRef.current = JSON.stringify(remote);
				setSlides(remote);
				return;
			}
		}
		const serialized = JSON.stringify(slides);
		if (serialized === lastSyncedRef.current) {
			return;
		}
		lastSyncedRef.current = serialized;

		void (async () => {
			const factories = await getFactories();
			if (failedRef.current) {
				return;
			}
			reconcileSlidesInYDoc(
				slides,
				doc as unknown as Parameters<typeof reconcileSlidesInYDoc>[1],
				factories,
				LOCAL_SYNC_ORIGIN,
			);
			scheduleWriteBack();
		})();
	}, [
		doc,
		slides,
		isConnected,
		isSynced,
		getFactories,
		scheduleWriteBack,
		readRemoteSlides,
		config?.role,
		setSlides,
	]);

	useEffect(() => {
		hasInitializedRef.current = false;
		lastSyncedRef.current = '';
	}, [doc]);

	useEffect(() => {
		if (!isConnected || !doc) {
			return;
		}

		const handleChange = (_events?: unknown, transaction?: YTransactionLike) => {
			if (transaction?.origin === LOCAL_SYNC_ORIGIN) {
				return;
			}
			const remoteSlides = readRemoteSlides(
				doc as unknown as Parameters<typeof readSlidesFromYDoc>[0],
			);
			if (remoteSlides.length === 0) {
				return;
			}

			const serialized = JSON.stringify(remoteSlides);
			if (serialized === lastSyncedRef.current) {
				return;
			}
			lastSyncedRef.current = serialized;

			isApplyingRemoteRef.current = true;
			setSlides(remoteSlides);
			isApplyingRemoteRef.current = false;
			scheduleWriteBack();
		};

		const unobserve = observeYDocSlides(
			doc as unknown as Parameters<typeof observeYDocSlides>[0],
			handleChange,
		);

		if (!hasInitializedRef.current) {
			hasInitializedRef.current = true;
			const arr = (doc as unknown as { getArray: (k: string) => { length: number } }).getArray(
				'pptx:slides',
			);
			if (arr.length > 0) {
				handleChange();
			}
		}

		return () => {
			unobserve();
			if (writeBackTimerRef.current !== null) {
				clearTimeout(writeBackTimerRef.current);
				writeBackTimerRef.current = null;
			}
		};
	}, [doc, isConnected, setSlides, scheduleWriteBack, readRemoteSlides]);
}
