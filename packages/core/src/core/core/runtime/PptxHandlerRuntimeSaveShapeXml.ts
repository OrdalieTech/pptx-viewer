import { XmlObject } from '../../types';
import type {
	ChartPptxElement,
	InkPptxElement,
	OlePptxElement,
	PptxElement,
	TablePptxElement,
} from '../../types';
import { resolveGroupChildBoxEmu, resolveGroupTightRewrap } from './group-tight-rewrap';
import type { GroupChildSpaceOwner } from './group-xfrm-preservation';
import type { SaveSlideContext } from './PptxHandlerRuntimeSaveElementEmbedding';
import { PptxHandlerRuntime as PptxHandlerRuntimeBase } from './PptxHandlerRuntimeSaveElements';
import type { SlideShapeCollectors } from './PptxHandlerRuntimeSaveElementWriter';
import {
	createGroupChildCollectors,
	pickGroupChildFromCollectors,
} from './save-group-child-collectors';
import { groupChildInheritedFill } from './save-group-fill';
import type { GroupChildEntry, GroupOwnEmuOverride } from './save-group-shape-xml';
import {
	appendGroupChildren,
	applyGroupChildTransform,
	buildGroupNonVisualXml,
	buildGroupPropertiesXml,
	buildGroupTransformXml,
	classifyGroupChildTag,
} from './save-group-shape-xml';
import { buildChartGraphicFrameXml, buildTableGraphicFrameXml } from './save-shape-xml-frames';
import { buildInkShapeXml } from './save-shape-xml-ink';
import {
	applyOleTypedFieldUpdatesXml,
	buildOleGraphicFrameXml,
	OLE_IMAGE_RELATIONSHIP_TYPE,
	OLE_OBJECT_RELATIONSHIP_TYPE,
	resolveOleEmbedRelationshipIdFromRels,
} from './save-shape-xml-ole';

/** Relationship type for chart parts. */
export const CHART_RELATIONSHIP_TYPE =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart';

/** Content type for a chart part in `[Content_Types].xml`. */
export const CHART_CONTENT_TYPE =
	'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';

/** The save context a group child needs to serialise like a top-level shape. */
export type GroupChildSaveContext = SaveSlideContext;

/**
 * Structural view of the element writer that lives further down the mixin
 * chain. `buildGroupShapeXml` is defined in an ancestor mixin, so
 * `processSlideElement` is present at runtime but not in this class's static
 * type; this interface names the one method needed without widening to `any`.
 */
interface GroupChildElementWriter {
	processSlideElement(
		el: PptxElement,
		collectors: SlideShapeCollectors,
		ctx: GroupChildSaveContext,
	): void;
}

export class PptxHandlerRuntime extends PptxHandlerRuntimeBase {
	/** See `save-shape-xml-frames.ts`'s `buildTableGraphicFrameXml`. */
	protected createTableGraphicFrameXml(el: TablePptxElement): XmlObject {
		return buildTableGraphicFrameXml(el, PptxHandlerRuntime.EMU_PER_PX);
	}

	/** See `save-shape-xml-frames.ts`'s `buildChartGraphicFrameXml`. */
	protected createChartGraphicFrameXml(
		el: ChartPptxElement,
		relId: string,
		extended = false,
	): XmlObject {
		return buildChartGraphicFrameXml(el, PptxHandlerRuntime.EMU_PER_PX, relId, extended);
	}

	/** See `save-shape-xml-ole.ts`'s `buildOleGraphicFrameXml`. */
	protected createOleGraphicFrameXml(el: OlePptxElement, embedRelationshipId: string): XmlObject {
		return buildOleGraphicFrameXml(el, PptxHandlerRuntime.EMU_PER_PX, embedRelationshipId);
	}

	/** See `save-shape-xml-ole.ts`'s `applyOleTypedFieldUpdatesXml`. */
	protected applyOleTypedFieldUpdates(shape: XmlObject, el: OlePptxElement): void {
		applyOleTypedFieldUpdatesXml(shape, el);
	}

	/** See `save-shape-xml-ole.ts`'s `resolveOleEmbedRelationshipIdFromRels`. */
	protected resolveOleEmbedRelationshipId(
		slideRelationships: XmlObject[],
		oleTarget: string | undefined,
	): string | undefined {
		return resolveOleEmbedRelationshipIdFromRels(slideRelationships, oleTarget);
	}

	/** Constants are exposed so the element-writer mixin can reuse them. */
	protected static readonly OLE_OBJECT_RELATIONSHIP_TYPE = OLE_OBJECT_RELATIONSHIP_TYPE;
	protected static readonly OLE_IMAGE_RELATIONSHIP_TYPE = OLE_IMAGE_RELATIONSHIP_TYPE;

	/** See `save-shape-xml-ink.ts`'s `buildInkShapeXml`. */
	protected createInkShapeXml(el: InkPptxElement): XmlObject {
		return buildInkShapeXml(el, PptxHandlerRuntime.EMU_PER_PX);
	}
}
