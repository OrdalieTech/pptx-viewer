import type { PptxChartData, XmlObject } from '../../types';
import { CHART_EX_REL_TYPE, CHART_REL_TYPE } from './chart-part-family-switch';
import {
	ensureContentTypeOverride,
	relativePartTarget,
	resolvePartTarget,
	rewriteChartRelationship,
} from './chart-part-registration';
import type { ChartPartRegistrationDeps } from './chart-part-registration';

const asArray = (value: unknown): XmlObject[] =>
	value === null || value === undefined
		? []
		: ((Array.isArray(value) ? value : [value]) as XmlObject[]);
/**
 * A slide copied from another slide shares the source's chart part (and the
 * part's workbook, style and colour parts). Before an edit is written, give
 * this slide its own copies, as PowerPoint does, so the edit stays on it.
 */
export async function forkSharedChartPart(
	deps: ChartPartRegistrationDeps,
	chartData: Pick<
		PptxChartData,
		'chartPartPath' | 'chartRelationshipId' | 'colorStylePartPath' | 'externalData'
	>,
	slidePath: string,
): Promise<void> {
	const oldPartPath = chartData.chartPartPath;
	const relsOf = (part: string) => part.replace(/([^/]+)$/, '_rels/$1.rels');
	const references = async (owner: string) => {
		const xml = await deps.zip.file(relsOf(owner))?.async('string');
		return xml
			? asArray(
					((deps.parser.parse(xml) as XmlObject)['Relationships'] as XmlObject | undefined)?.[
						'Relationship'
					],
				).filter(
					(rel) =>
						rel['@_TargetMode'] !== 'External' &&
						resolvePartTarget(owner, String(rel['@_Target'] ?? '')) === oldPartPath,
				).length
			: 0;
	};
	const slides = Object.keys(deps.zip.files).filter((path) =>
		/^ppt\/slides\/[^/]+\.xml$/.test(path),
	);
	if (!oldPartPath || !slides.includes(slidePath) || (await references(slidePath)) !== 1) {
		return;
	}
	let owners = 0;
	for (const slide of slides) {
		owners += (await references(slide)) > 0 ? 1 : 0;
	}
	if (owners < 2) {
		return;
	}
	const contentTypes = asArray(
		(
			(deps.parser.parse(await deps.zip.file('[Content_Types].xml')!.async('string')) as XmlObject)[
				'Types'
			] as XmlObject | undefined
		)?.['Override'],
	);
	const copies = /* @__PURE__ */ new Map<string, string>();
	const copyPart = async (path: string): Promise<string> => {
		const extension = /\.[^./]+$/.exec(path)?.[0] ?? '';
		const stem = path.slice(0, path.length - extension.length).replace(/\d+$/, '');
		let n = 1;
		while (deps.zip.file(`${stem}${n}${extension}`)) {
			n += 1;
		}
		const copy = `${stem}${n}${extension}`;
		copies.set(path, copy);
		deps.zip.file(copy, await deps.zip.file(path)!.async('uint8array'));
		const override = contentTypes.find((entry) => String(entry['@_PartName']) === `/${path}`);
		if (override) {
			await ensureContentTypeOverride(deps, copy, String(override['@_ContentType']));
		}
		const relsXml = await deps.zip.file(relsOf(path))?.async('string');
		if (relsXml) {
			const tree = deps.parser.parse(relsXml) as XmlObject;
			for (const rel of asArray(
				(tree['Relationships'] as XmlObject | undefined)?.['Relationship'],
			)) {
				const target = resolvePartTarget(path, String(rel['@_Target'] ?? ''));
				if (rel['@_TargetMode'] !== 'External' && deps.zip.file(target)) {
					rel['@_Target'] = relativePartTarget(
						copy,
						copies.get(target) ?? (await copyPart(target)),
					);
				}
			}
			deps.zip.file(relsOf(copy), deps.builder.build(tree));
		}
		return copy;
	};
	const newPartPath = await copyPart(oldPartPath);
	const relationshipId2 = await rewriteChartRelationship(deps, slidePath, {
		relationshipId: chartData.chartRelationshipId,
		oldPartPath,
		newPartPath,
		relationshipType: oldPartPath.startsWith('ppt/extendedCharts/')
			? CHART_EX_REL_TYPE
			: CHART_REL_TYPE,
	});
	chartData.chartPartPath = newPartPath;
	chartData.chartRelationshipId = relationshipId2 ?? chartData.chartRelationshipId;
	if (chartData.colorStylePartPath && copies.has(chartData.colorStylePartPath)) {
		chartData.colorStylePartPath = copies.get(chartData.colorStylePartPath);
	}
	if (chartData.externalData?.targetPath) {
		const workbook = copies.get(resolvePartTarget(oldPartPath, chartData.externalData.targetPath));
		if (workbook) {
			chartData.externalData = {
				...chartData.externalData,
				targetPath: relativePartTarget(newPartPath, workbook),
			};
		}
	}
}
/** PowerPoint expects one chart part per graphic frame: a slide that shares another slide's chart part gets a copy. */
export async function forkChartPartsSharedBetweenSlides(
	deps: ChartPartRegistrationDeps,
): Promise<void> {
	const number = (path: string) => Number(/(\d+)\.xml$/.exec(path)?.[1] ?? 0);
	const slides = Object.keys(deps.zip.files)
		.filter((path) => /^ppt\/slides\/slide\d+\.xml$/.test(path))
		.sort((a, b) => number(a) - number(b));
	const used = /* @__PURE__ */ new Set();
	for (const slidePath of slides) {
		const xml = await deps.zip
			.file(slidePath.replace(/([^/]+)$/, '_rels/$1.rels'))
			?.async('string');
		const relationships = xml
			? asArray(
					((deps.parser.parse(xml) as XmlObject)['Relationships'] as XmlObject | undefined)?.[
						'Relationship'
					],
				)
			: [];
		for (const rel of relationships) {
			if (
				rel['@_TargetMode'] === 'External' ||
				![CHART_REL_TYPE, CHART_EX_REL_TYPE].includes(String(rel['@_Type']))
			) {
				continue;
			}
			const part = resolvePartTarget(slidePath, String(rel['@_Target'] ?? ''));
			if (used.has(part)) {
				await forkSharedChartPart(
					deps,
					{ chartPartPath: part, chartRelationshipId: String(rel['@_Id']) },
					slidePath,
				);
			} else {
				used.add(part);
			}
		}
	}
}
