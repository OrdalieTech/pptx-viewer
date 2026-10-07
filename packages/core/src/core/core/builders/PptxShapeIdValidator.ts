import type { XmlObject } from '../../types';

/**
 * Shape ID uniqueness validator for OOXML slide shape trees.
 *
 * OpenXML requires that every `p:cNvPr/@id` within a single slide's
 * `p:spTree` is unique. Duplicate IDs can corrupt files in MS Office.
 * This validator scans the tree and reassigns duplicate IDs.
 */

/** Recursively collect all cNvPr nodes from a shape tree. */
function collectCnvPrNodes(
	node: XmlObject,
	results: XmlObject[],
	ensureArray: (value: unknown) => unknown[],
): void {
	// Check direct cNvPr references in nvSpPr, nvPicPr, nvCxnSpPr, nvGrpSpPr, nvGraphicFramePr
	const nvContainers = [
		'p:nvSpPr',
		'p:nvPicPr',
		'p:nvCxnSpPr',
		'p:nvGrpSpPr',
		'p:nvGraphicFramePr',
		'p:nvContentPartPr',
	];
	for (const nvKey of nvContainers) {
		const nvNode = node[nvKey] as XmlObject | undefined;
		if (nvNode?.['p:cNvPr']) {
			results.push(nvNode['p:cNvPr'] as XmlObject);
		}
	}
	// A real `p:contentPart`'s non-visual properties are `p14:`-qualified, not
	// `p:`-qualified (verified against PowerPoint's own SaveAs output; see
	// `mc-capabilities.ts`), so a content part's id lives at
	// `p14:nvContentPartPr/p14:cNvPr`. Missing this left the id invisible to
	// this validator: it could no longer detect (or dedupe) a collision
	// between an authored content part and an ordinary shape, so a freshly
	// drawn stroke could silently reuse another shape's id and produce a
	// package PowerPoint's own reader rejects as corrupted.
	const p14ContentPartNv = node['p14:nvContentPartPr'] as XmlObject | undefined;
	if (p14ContentPartNv?.['p14:cNvPr']) {
		results.push(p14ContentPartNv['p14:cNvPr'] as XmlObject);
	}

	// Recurse into shape lists
	const shapeLists = ['p:sp', 'p:pic', 'p:cxnSp', 'p:graphicFrame', 'p:grpSp', 'p:contentPart'];
	for (const listKey of shapeLists) {
		const children = ensureArray(node[listKey]) as XmlObject[];
		for (const child of children) {
			collectCnvPrNodes(child, results, ensureArray);
		}
	}

	// Ink and other Office extensions place their real element and fallback
	// shape inside mc:AlternateContent branches. Those nodes still occupy the
	// slide's non-visual ID space and must participate in duplicate
	// detection, or a Draw operation can introduce a repeated id that makes
	// desktop PowerPoint repair or reject the deck.
	for (const alternate of ensureArray(node['mc:AlternateContent']) as XmlObject[]) {
		for (const branchKey of ['mc:Choice', 'mc:Fallback']) {
			for (const branch of ensureArray(alternate[branchKey]) as XmlObject[]) {
				collectCnvPrNodes(branch, results, ensureArray);
			}
		}
	}
}

export interface IPptxShapeIdValidator {
	validateAndDeduplicateIds(spTree: XmlObject, ensureArray: (value: unknown) => unknown[]): number;
}

/**
 * Validates shape IDs in a slide's spTree and reassigns duplicates.
 * Returns the number of IDs that were reassigned.
 */
export class PptxShapeIdValidator implements IPptxShapeIdValidator {
	public validateAndDeduplicateIds(
		spTree: XmlObject,
		ensureArray: (value: unknown) => unknown[],
	): number {
		const cNvPrNodes: XmlObject[] = [];
		collectCnvPrNodes(spTree, cNvPrNodes, ensureArray);

		if (cNvPrNodes.length === 0) {
			return 0;
		}

		const shapeLists = ['p:sp', 'p:pic', 'p:cxnSp', 'p:graphicFrame', 'p:grpSp', 'p:contentPart'];
		const maxId = cNvPrNodes.reduce(
			(max, node) => Math.max(max, Number.parseInt(String(node['@_id']), 10) || 0),
			0,
		);
		let reassigned = 0;
		let nextId = maxId;

		const directCnvPrNodes = (node: XmlObject): XmlObject[] => {
			const result: XmlObject[] = [];
			for (const key of [
				'p:nvSpPr',
				'p:nvPicPr',
				'p:nvCxnSpPr',
				'p:nvGrpSpPr',
				'p:nvGraphicFramePr',
				'p:nvContentPartPr',
			]) {
				const cNvPr = (node[key] as XmlObject | undefined)?.['p:cNvPr'];
				if (cNvPr) result.push(cNvPr as XmlObject);
			}
			const p14CnvPr = (node['p14:nvContentPartPr'] as XmlObject | undefined)?.['p14:cNvPr'];
			if (p14CnvPr) result.push(p14CnvPr as XmlObject);
			return result;
		};

		const visit = (node: XmlObject, usedIds: Set<number>): void => {
			for (const cNvPr of directCnvPrNodes(node)) {
				let id = Number.parseInt(String(cNvPr['@_id'] ?? '0'), 10) || 0;
				if (id <= 0 || usedIds.has(id)) {
					id = ++nextId;
					cNvPr['@_id'] = String(id);
					reassigned++;
				}
				usedIds.add(id);
			}

			for (const key of shapeLists) {
				for (const child of ensureArray(node[key]) as XmlObject[]) visit(child, usedIds);
			}

			// Choice and Fallback are alternatives in OOXML. Their shapes may
			// legitimately carry matching IDs, while collisions with the base
			// slide or within one branch must still be repaired.
			for (const alternate of ensureArray(node['mc:AlternateContent']) as XmlObject[]) {
				const branchIds: Set<number>[] = [];
				for (const key of ['mc:Choice', 'mc:Fallback']) {
					for (const branch of ensureArray(alternate[key]) as XmlObject[]) {
						const branchUsedIds = new Set(usedIds);
						visit(branch, branchUsedIds);
						branchIds.push(branchUsedIds);
					}
				}
				for (const ids of branchIds) {
					for (const id of ids) usedIds.add(id);
				}
			}
		};

		visit(spTree, new Set());

		return reassigned;
	}
}
