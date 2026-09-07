import type { YTextLike } from './collaboration-text-codec';

export interface YMapLike {
	get: (key: string) => unknown;
	set: (key: string, value: unknown) => void;
	delete: (key: string) => void;
	forEach: (cb: (value: unknown, key: string) => void) => void;
}

/** Shape of the Yjs transaction passed to (deep) observers. */
export interface YTransactionLike {
	origin?: unknown;
}

export type YDeepObserver = (events?: unknown, transaction?: YTransactionLike) => void;

export interface YArrayLike {
	readonly length: number;
	get: (index: number) => unknown;
	push: (items: unknown[]) => void;
	delete: (index: number, length?: number) => void;
	insert: (index: number, items: unknown[]) => void;
	toArray: () => unknown[];
	observe: (handler: () => void) => void;
	unobserve: (handler: () => void) => void;
	observeDeep: (handler: YDeepObserver) => void;
	unobserveDeep: (handler: YDeepObserver) => void;
}

export interface YDocLike {
	getMap: (name: string) => YMapLike;
	getArray: (name: string) => YArrayLike;
	transact: (fn: () => void, origin?: unknown) => void;
}

export interface YjsFactories {
	createMap: () => YMapLike;
	createArray: () => YArrayLike;
	createText: () => YTextLike;
}

// ---------------------------------------------------------------------------
// Y.Doc schema constants
// ---------------------------------------------------------------------------

export const YDOC_SLIDES_KEY = 'pptx:slides';
export const YDOC_META_KEY = 'pptx:meta';
export const YDOC_SCHEMA_VERSION = 1;

export function assertCollaborationSchema(ydoc: YDocLike): void {
	const version = ydoc.getMap(YDOC_META_KEY).get('schemaVersion');
	if (version !== undefined && version !== YDOC_SCHEMA_VERSION) {
		throw new Error(`Unsupported PPTX collaboration schema: ${String(version)}`);
	}
	if (ydoc.getMap(YDOC_META_KEY).get('sourceBytes') !== undefined) {
		throw new Error('Legacy PPTX state contains a source binary; explicit migration required');
	}
}

export const SCALAR_ELEMENT_KEYS: ReadonlySet<string> = new Set([
	'id',
	'type',
	'x',
	'y',
	'width',
	'height',
	'rotation',
	'shapeId',
	'skewX',
	'skewY',
	'flipHorizontal',
	'flipVertical',
	'hidden',
	'opacity',
	'text',
	'name',
	'altText',
	'shapeType',
	'imagePath',
	'imageData',
	'svgData',
	'svgPath',
	'cropLeft',
	'cropTop',
	'cropRight',
	'cropBottom',
	'tileOffsetX',
	'tileOffsetY',
	'tileScaleX',
	'tileScaleY',
	'tileFlip',
	'tileAlignment',
	'pathData',
	'pathWidth',
	'pathHeight',
	'mediaType',
	'mediaPath',
	'mediaMimeType',
	'mediaReferenceKind',
	'mediaReferenceName',
	'mediaReferenceContentType',
	'trimStartMs',
	'trimEndMs',
	'posterFramePath',
	'fullScreen',
	'loop',
	'fadeInDuration',
	'fadeOutDuration',
	'volume',
	'autoPlay',
	'playAcrossSlides',
	'hideWhenNotPlaying',
	'playbackSpeed',
	'mediaMissing',
	'isLinked',
	'oleTarget',
	'oleProgId',
	'oleName',
	'oleClsId',
	'oleObjectType',
	'oleFileExtension',
	'fileName',
	'externalPath',
	'previewImage',
	'oleShowAsIcon',
	'oleImgW',
	'oleImgH',
	'oleEmbeddedFileName',
	'oleEmbeddedMimeType',
	'oleEmbeddedByteSize',
	'inkPaths',
	'inkColors',
	'inkWidths',
	'inkOpacities',
	'inkTool',
	'inkPartPath',
	'zoomType',
	'targetSlideIndex',
	'targetSectionId',
	'summaryLayout',
	'modelPath',
	'modelMimeType',
	'posterImage',
	'linkedTxbxId',
	'linkedTxbxSeq',
	'promptText',
]);

export const COMPLEX_ELEMENT_FIELDS: Readonly<Record<string, string>> = {
	textStyle: '_ts',
	shapeStyle: '_ss',
	shapeAdjustments: '_sa',
	adjustmentHandles: '_ah',
	chartData: '_cd',
	smartArtData: '_smad',
	paragraphIndents: '_pi',
	rawXml: '_rx',
	extLstXml: '_elx',
	actionClick: '_ac',
	actionHover: '_av',
	locks: '_lk',
	imageEffects: '_ie',
	cropShape: '_cr',
	bookmarks: '_mb',
	captionTracks: '_ct',
	audioCdStart: '_acd1',
	audioCdEnd: '_acd2',
	rawMediaReferenceXml: '_mrx',
	metadata: '_md',
	groupFill: '_gf',
	inkPointPressures: '_ipp',
	inkStrokes: '_cis',
	inkPartRawXml: '_cirx',
	summaryTargets: '_zst',
	extensionXml: '_ext',
	customGeometryPaths: '_cgp',
	customGeometryRawData: '_cgr',
	customGeometryAdjustHandlesXY: '_cgx',
	customGeometryAdjustHandlesPolar: '_cgo',
	customGeometryConnectionSites: '_cgc',
	customGeometryTextRect: '_cgt',
};
export const SCALAR_SLIDE_KEYS: ReadonlySet<string> = new Set([
	'id',
	'rId',
	'sourceSlideId',
	'name',
	'layoutPath',
	'layoutName',
	'slideNumber',
	'hidden',
	'sectionName',
	'sectionId',
	'backgroundColor',
	'backgroundImage',
	'backgroundGradient',
	'backgroundShadeToTitle',
	'notes',
	'notesCSldName',
	'backgroundShowAnimation',
	'showMasterShapes',
	'isDirty',
]);

export const COMPLEX_SLIDE_FIELDS: Readonly<Record<string, string>> = {
	transition: '_tr',
	animations: '_an',
	nativeAnimations: '_na',
	rawTiming: '_rt',
	notesSegments: '_ns',
	notesShapes: '_nsh',
	notesClrMapOverride: '_ncm',
	comments: '_cm',
	warnings: '_wa',
	rawXml: '_rx',
	clrMapOverride: '_cm2',
	guides: '_gu',
	customerData: '_cu',
	activeXControls: '_ax',
	legacyVmlElements: '_lvml',
	backgroundPattern: '_bp',
	modernCommentPart: '_mc',
	headerFooterFlags: '_hff',
	slideSynchronization: '_sync',
};
