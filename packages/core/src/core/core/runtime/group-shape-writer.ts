import type { GroupPptxElement, PptxElement, XmlObject } from '../../types';

const loadedGroups = new WeakMap<XmlObject, string>();
const loadedTemplates = new WeakMap<object, Map<string, string>>();

function groupState(value: unknown): string {
	return JSON.stringify(value, (key, entry) => {
		if (key === 'rawXml' || key === 'id' || key === 'shapeId') {
			return undefined;
		}
		if (entry && typeof entry === 'object' && !Array.isArray(entry)) {
			return Object.fromEntries(
				Object.keys(entry)
					.sort()
					.map((field) => [field, entry[field]]),
			);
		}
		return entry;
	});
}

export function rememberLoadedGroup(group: GroupPptxElement): void {
	if (group.rawXml) {
		loadedGroups.set(group.rawXml, groupState(group));
	}
}

export function rememberLoadedShapes(owner: object, elements: PptxElement[]): void {
	const templates = new Map<string, string>();
	const visit = (items: PptxElement[]): void => {
		for (const element of items) {
			templates.set(element.id, groupState(element));
			if (element.type === 'group') {
				rememberLoadedGroup(element);
				visit(element.children);
			}
		}
	};
	visit(elements);
	loadedTemplates.set(owner, templates);
}

export function isUnchangedTemplate(owner: object, element: PptxElement): boolean {
	return loadedTemplates.get(owner)?.get(element.id) === groupState(element);
}

export function findOriginalGroup(
	tree: XmlObject | undefined,
	raw: XmlObject | undefined,
): XmlObject | undefined {
	const id = (raw?.['p:nvGrpSpPr'] as XmlObject | undefined)?.['p:cNvPr'] as XmlObject | undefined;
	if (!tree || id?.['@_id'] === undefined) {
		return undefined;
	}
	const groups = tree['p:grpSp'];
	for (const group of (Array.isArray(groups) ? groups : groups ? [groups] : []) as XmlObject[]) {
		const properties = (group['p:nvGrpSpPr'] as XmlObject | undefined)?.['p:cNvPr'] as
			| XmlObject
			| undefined;
		if (String(properties?.['@_id']) === String(id['@_id'])) {
			return group;
		}
		const nested = findOriginalGroup(group, raw);
		if (nested) {
			return nested;
		}
	}
	return undefined;
}

/** Preserve untouched native group structure; edited children use the normal element writer. */
export function writeGroupShape(
	group: GroupPptxElement,
	emu: number,
	writeChildren: (children: PptxElement[]) => Record<string, XmlObject[]>,
	original?: XmlObject,
): XmlObject {
	const source = original ?? group.rawXml;
	if (source && loadedGroups.get(source) === groupState(group)) {
		return source;
	}
	const xml: XmlObject = group.rawXml
		? structuredClone(group.rawXml)
		: {
				'p:nvGrpSpPr': {
					'p:cNvPr': { '@_id': group.shapeId ?? '0', '@_name': group.name ?? group.id },
					'p:cNvGrpSpPr': {},
					'p:nvPr': {},
				},
			};
	const properties = (xml['p:grpSpPr'] ?? {}) as XmlObject;
	const transform = (properties['a:xfrm'] ?? {}) as XmlObject;
	transform['a:off'] = {
		'@_x': String(Math.round(group.x * emu)),
		'@_y': String(Math.round(group.y * emu)),
	};
	transform['a:ext'] = {
		'@_cx': String(Math.round(group.width * emu)),
		'@_cy': String(Math.round(group.height * emu)),
	};
	// Typed children already use rendered pixels relative to their parent.
	transform['a:chOff'] = { '@_x': '0', '@_y': '0' };
	transform['a:chExt'] = structuredClone(transform['a:ext']);
	if (group.rotation !== undefined) {
		transform['@_rot'] = String(Math.round(group.rotation * 60000));
	}
	for (const [field, flag] of [
		['@_flipH', group.flipHorizontal],
		['@_flipV', group.flipVertical],
	] as const) {
		if (flag) {
			transform[field] = '1';
		} else {
			delete transform[field];
		}
	}
	properties['a:xfrm'] = transform;
	xml['p:grpSpPr'] = properties;
	for (const [tag, children] of Object.entries(writeChildren(group.children))) {
		if (children.length) {
			xml[tag] = children;
		} else {
			delete xml[tag];
		}
	}
	return xml;
}
