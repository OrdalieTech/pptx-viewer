export {
	PptxCodec,
	yjsFactories,
	ORIGIN_FILE_LOAD,
	SCALAR_ELEMENT_KEYS,
	COMPLEX_FIELD_MAP,
	SCALAR_SLIDE_KEYS,
	COMPLEX_SLIDE_FIELD_MAP,
} from './pptx-codec.js';
export type { FormatCodec } from './pptx-codec.js';
export {
	readSlidesFromYDoc,
	writeSlidesToYDoc,
	writeSlideToYMap,
	writeElementToYMap,
	readElementFromYMap,
	reconcileSlidesInYDoc,
	findElementYMap,
	registerCollaborationSource,
	assertCollaborationSchema,
	YDOC_SCHEMA_VERSION,
	COMPLEX_ELEMENT_FIELDS,
	COMPLEX_SLIDE_FIELDS,
	orderedYMaps,
	reorderYArray,
	writeTableData,
	reconcileTableData,
	readTableData,
} from 'pptx-viewer-shared/collaboration';
export type {
	YDocLike,
	YMapLike,
	YArrayLike,
	YjsFactories,
} from 'pptx-viewer-shared/collaboration';
