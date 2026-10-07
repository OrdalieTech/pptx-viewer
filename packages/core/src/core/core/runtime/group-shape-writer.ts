import type { PptxElement, XmlObject } from '../../types';

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

export function rememberLoadedShapes(owner: object, elements: PptxElement[]): void {
	const templates = new Map<string, string>();
	const visit = (items: PptxElement[]): void => {
		for (const element of items) {
			templates.set(element.id, groupState(element));
			if (element.type === 'group') {
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

/**
 * Replace the load-time state of these elements with the state a collaboration
 * store reads back for them, so an element that went through that store
 * unchanged still counts as unchanged (and keeps its original XML).
 */
export function rememberCollaborationBaseline(owner: object, elements2: PptxElement[]): void {
	const templates = loadedTemplates.get(owner) ?? /* @__PURE__ */ new Map<string, string>();
	const visit = (items: PptxElement[]): void => {
		for (const element of items) {
			templates.set(element.id, groupState(element));
			if (element.type === 'group') {
				visit(element.children);
			}
		}
	};
	visit(elements2);
	loadedTemplates.set(owner, templates);
}

export function loadedTextSegments(
	owner: object,
	id: string,
): import('../../types').TextSegment[] | undefined {
	const state = loadedTemplates.get(owner)?.get(id);
	return state ? JSON.parse(state).textSegments : undefined;
}
