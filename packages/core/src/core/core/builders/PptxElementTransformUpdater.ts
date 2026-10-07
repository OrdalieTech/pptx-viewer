import type { PptxElement, XmlObject } from '../../types';
import { isXfrmEmuUnchanged, resolveXfrmEmu } from '../../utils/xfrm-emu-resolution';
import { resolveGroupChildBoxEmu } from '../runtime/group-tight-rewrap';
import type { GroupChildSpaceOwner } from '../runtime/group-xfrm-preservation';

export interface IPptxElementTransformUpdater {
	applyTransform(
		shape: XmlObject,
		element: PptxElement,
		emuPerPx: number,
		enclosingGroupChildSpace?: GroupChildSpaceOwner,
	): void;
}

export class PptxElementTransformUpdater implements IPptxElementTransformUpdater {
	public applyTransform(
		shape: XmlObject,
		element: PptxElement,
		emuPerPx: number,
		enclosingGroupChildSpace?: GroupChildSpaceOwner,
	): void {
		let transform = ((shape['p:spPr'] as XmlObject | undefined)?.['a:xfrm'] || shape['p:xfrm']) as
			| XmlObject
			| undefined;
		if (!transform) {
			const edited =
				(
					[
						[element.x, element.xEmu],
						[element.y, element.yEmu],
						[element.width, element.widthEmu],
						[element.height, element.heightEmu],
					] as const
				).some(([px, emu]) => !isXfrmEmuUnchanged(px, emu, emuPerPx)) ||
				Boolean(element.rotation) ||
				Boolean(element.flipHorizontal) ||
				Boolean(element.flipVertical);
			if (!edited || enclosingGroupChildSpace || !('p:spPr' in shape)) {
				return;
			}
			const spPr =
				shape['p:spPr'] && typeof shape['p:spPr'] === 'object'
					? (shape['p:spPr'] as XmlObject)
					: {};
			transform = {};
			shape['p:spPr'] = {
				...Object.fromEntries(Object.entries(spPr).filter(([key]) => key.startsWith('@_'))),
				'a:xfrm': transform,
				...spPr,
			};
		}

		if (!transform['a:off']) {
			transform['a:off'] = {};
		}
		if (!transform['a:ext']) {
			transform['a:ext'] = {};
		}

		// `resolveXfrmEmu` re-emits the exact EMU an untouched axis was parsed
		// from (byte-identical `a:off`/`a:ext`) instead of re-quantizing from
		// pixels, but only when the element still reports the same pixel value
		// it was parsed with; once the model's x/y/width/height diverges (the
		// element was moved or resized) it falls back to `px * emuPerPx`
		// exactly as before. See `xfrm-emu-resolution.ts`.
		//
		// `enclosingGroupChildSpace` (set only for a direct child of a group;
		// see `group-xfrm-preservation.ts`) inverts this element's CURRENT
		// relative-to-group pixel geometry back into that group's captured
		// child space instead: its exact original child-space EMU verbatim
		// when unchanged (bypassing `resolveXfrmEmu`'s parent-space comparison,
		// which can spuriously disagree even when nothing moved - see the
		// module doc in `group-xfrm-preservation.ts`), or the inverse of the
		// parse-time mapping when it has moved/resized, or the group itself has.
		// `resolveGroupChildBoxEmu` (`group-tight-rewrap.ts`) additionally
		// recurses into `element` first when it is itself a group, so ITS OWN
		// re-wrapped box (not its possibly-stale relative geometry) is what
		// gets inverted into the enclosing group's space.
		const inverted = enclosingGroupChildSpace
			? resolveGroupChildBoxEmu(element, enclosingGroupChildSpace, emuPerPx)
			: undefined;
		const extCx = inverted
			? inverted.widthEmu
			: resolveXfrmEmu(element.width, element.widthEmu, emuPerPx);
		const extCy = inverted
			? inverted.heightEmu
			: resolveXfrmEmu(element.height, element.heightEmu, emuPerPx);

		let offX: number;
		let offY: number;
		if (inverted) {
			offX = inverted.xEmu;
			offY = inverted.yEmu;
		} else {
			// Model coordinates already describe the unrotated frame used by a:off.
			offX = resolveXfrmEmu(element.x, element.xEmu, emuPerPx);
			offY = resolveXfrmEmu(element.y, element.yEmu, emuPerPx);
		}

		(transform['a:off'] as XmlObject)['@_x'] = String(offX);
		(transform['a:off'] as XmlObject)['@_y'] = String(offY);
		(transform['a:ext'] as XmlObject)['@_cx'] = String(extCx);
		(transform['a:ext'] as XmlObject)['@_cy'] = String(extCy);

		if (element.rotation !== undefined) {
			transform['@_rot'] = String(Math.round(element.rotation * 60000));
		}
		if (element.skewX !== undefined) {
			transform['@_skewX'] = String(Math.round(element.skewX * 60000));
		}
		if (element.skewY !== undefined) {
			transform['@_skewY'] = String(Math.round(element.skewY * 60000));
		}
		if (element.flipHorizontal) {
			transform['@_flipH'] = '1';
		} else {
			delete transform['@_flipH'];
		}
		if (element.flipVertical) {
			transform['@_flipV'] = '1';
		} else {
			delete transform['@_flipV'];
		}
	}
}
