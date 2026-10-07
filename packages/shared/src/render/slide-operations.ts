/**
 * slide-operations: Pure factory helpers for blank slides + slide ids.
 *
 * Framework-agnostic (no framework imports). The React `useSlideManagement`
 * hook and the Vue `useSlideOperations` composable both call these so the
 * blank-slide shape and id scheme live in one place.
 *
 * `makeSlideId` uses the `slide-${Date.now()}-${random}` pattern inside its
 * function body (called at runtime, never at module-eval time) and accepts an
 * optional `idGenerator` override, matching the neighbouring shared editor
 * modules' id convention.
 */

import type { PptxSlide } from 'pptx-viewer-core';

/**
 * Generate a collision-resistant slide id of the form
 * `slide-<timestamp>-<base36 suffix>`.
 *
 * @param idGenerator - Optional override; when supplied its result is returned
 *   verbatim (callers wanting `crypto.randomUUID` ids pass it here).
 */
export function makeSlideId(idGenerator?: () => string): string {
	if (idGenerator) {
		return idGenerator();
	}
	return `slide-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`;
}

/**
 * Build a fresh, empty {@link PptxSlide}: a new id, empty `rId`, the given
 * `slideNumber`, and an empty `elements` array. This is the minimal shape both
 * the React `handleAddSlide` and the Vue `createBlankSlide` produced.
 *
 * @param slideNumber - 1-based slide number for the new slide.
 * @param idGenerator - Optional id override forwarded to {@link makeSlideId}.
 */
export function createBlankSlide(
	slideNumber: number,
	idGenerator?: () => string,
	reference?: PptxSlide,
): PptxSlide {
	return {
		id: makeSlideId(idGenerator),
		rId: '',
		slideNumber,
		elements: [],
		// ponytail: a new slide uses its neighbour's layout and resolved background.
		layoutPath: reference?.layoutPath,
		showMasterShapes: reference?.showMasterShapes,
		backgroundColor: reference?.backgroundColor,
		backgroundGradient: reference?.backgroundGradient,
		backgroundImage: reference?.backgroundImage,
		backgroundImageProperties: reference?.backgroundImageProperties,
		backgroundPattern: reference?.backgroundPattern,
	};
}
