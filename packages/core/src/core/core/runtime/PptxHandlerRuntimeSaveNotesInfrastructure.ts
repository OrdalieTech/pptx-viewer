import type { PptxSlide, XmlObject } from '../../types';
import type { PptxSaveConstants } from '../factories';
import { relativePartTarget, resolvePartTarget } from './chart-part-registration';
import { PptxHandlerRuntime as PptxHandlerRuntimeBase } from './PptxHandlerRuntimeSaveViewProperties';

const NOTES_MASTER_CONTENT_TYPE =
	'application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml';

function asArray(value: unknown): XmlObject[] {
	return Array.isArray(value) ? (value as XmlObject[]) : value ? [value as XmlObject] : [];
}

function relationshipBase(strict: boolean): string {
	return strict
		? 'http://purl.oclc.org/ooxml/officeDocument/relationships/'
		: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
}

function presentationNamespace(strict: boolean): string {
	return strict
		? 'http://purl.oclc.org/ooxml/presentationml/main'
		: 'http://schemas.openxmlformats.org/presentationml/2006/main';
}

function drawingNamespace(strict: boolean): string {
	return strict
		? 'http://purl.oclc.org/ooxml/drawingml/main'
		: 'http://schemas.openxmlformats.org/drawingml/2006/main';
}

export class PptxHandlerRuntime extends PptxHandlerRuntimeBase {
	/** Ensure authored notes have the global notes-master OPC infrastructure. */
	protected async ensureNotesMasterForAuthoredNotes(
		slides: readonly PptxSlide[],
		constants: PptxSaveConstants,
	): Promise<void> {
		const needsNotes = slides.some(
			(slide) => slide.notes !== undefined || Boolean(slide.notesSegments?.length),
		);
		if (
			!needsNotes ||
			Object.keys(this.zip.files).some((path) =>
				/^ppt\/notesMasters\/notesMaster\d+\.xml$/u.test(path),
			)
		) {
			return;
		}

		const strict = constants.conformance === 'strict';
		const masterPath = 'ppt/notesMasters/notesMaster1.xml';
		const themePath =
			Object.keys(this.zip.files).find((path) => /^ppt\/theme\/theme\d+\.xml$/u.test(path)) ??
			'ppt/theme/theme1.xml';
		this.zip.file(masterPath, this.builder.build(this.buildDefaultNotesMaster(strict)));
		this.zip.file(
			'ppt/notesMasters/_rels/notesMaster1.xml.rels',
			this.builder.build({
				Relationships: {
					'@_xmlns': constants.relationshipsNamespace,
					Relationship: {
						'@_Id': 'rId1',
						'@_Type': `${relationshipBase(strict)}theme`,
						'@_Target': `../theme/${themePath.slice(themePath.lastIndexOf('/') + 1)}`,
					},
				},
			}),
		);

		const relsPath = 'ppt/_rels/presentation.xml.rels';
		const relsXml = await this.zip.file(relsPath)?.async('string');
		if (!relsXml || !this.presentationData) {
			return;
		}
		const relsData = this.parser.parse(relsXml) as XmlObject;
		const relsRoot = (relsData['Relationships'] ?? {}) as XmlObject;
		const relationships = asArray(relsRoot['Relationship']);
		const usedIds = new Set(relationships.map((rel) => String(rel['@_Id'] ?? '')));
		let index = 1;
		while (usedIds.has(`rId${index}`)) {
			index += 1;
		}
		const relationshipId = `rId${index}`;
		relationships.push({
			'@_Id': relationshipId,
			'@_Type': `${relationshipBase(strict)}notesMaster`,
			'@_Target': 'notesMasters/notesMaster1.xml',
		});
		relsRoot['Relationship'] = relationships;
		relsData['Relationships'] = relsRoot;
		this.zip.file(relsPath, this.builder.build(relsData));

		const presentation = this.presentationData['p:presentation'] as XmlObject;
		const notesMasterIdList: XmlObject = {
			'p:notesMasterId': { '@_r:id': relationshipId },
		};
		const reordered: XmlObject = {};
		let inserted = false;
		for (const [key, value] of Object.entries(presentation)) {
			if (!inserted && (key === 'p:handoutMasterIdLst' || key === 'p:sldIdLst')) {
				reordered['p:notesMasterIdLst'] = notesMasterIdList;
				inserted = true;
			}
			reordered[key] = value;
		}
		if (!inserted) {
			reordered['p:notesMasterIdLst'] = notesMasterIdList;
		}
		for (const key of Object.keys(presentation)) {
			delete presentation[key];
		}
		Object.assign(presentation, reordered);
		await this.addContentTypeOverride(`/${masterPath}`, NOTES_MASTER_CONTENT_TYPE);
	}

	// PowerPoint repairs notes masters that share a theme with slide/presentation parts.
	// ponytail: copy only shared notes themes, keeping their bytes and relative asset links.
	protected async isolateNotesMasterThemes(): Promise<void> {
		// OPC part names are ASCII case-insensitive; keep the ZIP spelling for reads.
		const parts = new Map(
			Object.keys(this.zip.files)
				.filter((path) => !this.zip.files[path].dir)
				.map((path) => [path.toLowerCase(), path]),
		);
		const themes: {
			owner: string;
			relsPath: string;
			tree: XmlObject;
			rel: XmlObject;
			target: string;
		}[] = [];
		for (const relsPath of Object.keys(this.zip.files).filter((path) => path.endsWith('.rels'))) {
			const owner = relsPath.replace(/(^|\/)_rels\//u, '$1').slice(0, -5);
			const tree = this.parser.parse(await this.zip.file(relsPath)!.async('string')) as XmlObject;
			for (const rel of asArray(
				(tree['Relationships'] as XmlObject | undefined)?.['Relationship'],
			)) {
				if (String(rel['@_Type'] ?? '').endsWith('/theme') && rel['@_TargetMode'] !== 'External') {
					themes.push({
						owner,
						relsPath,
						tree,
						rel,
						target: resolvePartTarget(owner, String(rel['@_Target'] ?? '')).toLowerCase(),
					});
				}
			}
		}
		const isNotes = (entry: { owner: string }) =>
			entry.owner.toLowerCase().startsWith('ppt/notesmasters/');
		const used = new Set(themes.filter((entry) => !isNotes(entry)).map((entry) => entry.target));
		for (const entry of themes.filter(isNotes)) {
			const source = parts.get(entry.target);
			if (used.has(entry.target) && source) {
				const folder = source.slice(0, source.lastIndexOf('/') + 1);
				let index = 1;
				while (parts.has(`${folder}theme${index}.xml`.toLowerCase())) {
					index += 1;
				}
				const copy = `${folder}theme${index}.xml`;
				this.zip.file(copy, await this.zip.file(source)!.async('uint8array'));
				parts.set(copy.toLowerCase(), copy);
				const relsOf = (path: string) => path.replace(/([^/]+)$/u, '_rels/$1.rels');
				const assetsPath = parts.get(relsOf(source).toLowerCase());
				const assets = assetsPath ? this.zip.file(assetsPath) : null;
				if (assets) {
					this.zip.file(relsOf(copy), await assets.async('uint8array'));
				}
				await this.addContentTypeOverride(
					`/${copy}`,
					'application/vnd.openxmlformats-officedocument.theme+xml',
				);
				entry.rel['@_Target'] = relativePartTarget(entry.owner, copy);
				this.zip.file(entry.relsPath, this.builder.build(entry.tree));
				entry.target = copy.toLowerCase();
			}
			used.add(entry.target);
		}
	}

	private buildDefaultNotesMaster(strict: boolean): XmlObject {
		return {
			'p:notesMaster': {
				'@_xmlns:a': drawingNamespace(strict),
				'@_xmlns:r': `${relationshipBase(strict).slice(0, -1)}`,
				'@_xmlns:p': presentationNamespace(strict),
				'p:cSld': {
					'p:spTree': {
						'p:nvGrpSpPr': {
							'p:cNvPr': { '@_id': '1', '@_name': '' },
							'p:cNvGrpSpPr': {},
							'p:nvPr': {},
						},
						'p:grpSpPr': {
							'a:xfrm': {
								'a:off': { '@_x': '0', '@_y': '0' },
								'a:ext': { '@_cx': '0', '@_cy': '0' },
								'a:chOff': { '@_x': '0', '@_y': '0' },
								'a:chExt': { '@_cx': '0', '@_cy': '0' },
							},
						},
					},
				},
				'p:clrMap': {
					'@_accent1': 'accent1',
					'@_accent2': 'accent2',
					'@_accent3': 'accent3',
					'@_accent4': 'accent4',
					'@_accent5': 'accent5',
					'@_accent6': 'accent6',
					'@_bg1': 'lt1',
					'@_bg2': 'lt2',
					'@_folHlink': 'folHlink',
					'@_hlink': 'hlink',
					'@_tx1': 'dk1',
					'@_tx2': 'dk2',
				},
			},
		};
	}

	private async addContentTypeOverride(partName: string, contentType: string): Promise<void> {
		const xml = await this.zip.file('[Content_Types].xml')?.async('string');
		if (!xml) {
			return;
		}
		const data = this.parser.parse(xml) as XmlObject;
		const root = (data['Types'] ?? {}) as XmlObject;
		const overrides = asArray(root['Override']);
		if (!overrides.some((entry) => entry['@_PartName'] === partName)) {
			overrides.push({ '@_PartName': partName, '@_ContentType': contentType });
		}
		root['Override'] = overrides;
		data['Types'] = root;
		this.zip.file('[Content_Types].xml', this.builder.build(data));
	}
}
