import type { XMLBuilder, XMLParser } from 'fast-xml-parser';
import type JSZip from 'jszip';

import type { CompatibilityWarningInput } from '../../services';
import type { PptxSlide, XmlObject } from '../../types';
import { relativePartTarget, resolvePartTarget } from '../runtime/chart-part-registration';
import { buildNotesSlide } from './notes-slide-xml';
import type { IPptxSlideRelationshipRegistry } from './PptxSlideRelationshipRegistry';

export interface PptxSlideNotesPartUpdaterInput {
	slide: PptxSlide;
	relationshipRegistry: IPptxSlideRelationshipRegistry;
	slideNotesRelationshipType: string;
	zip: JSZip;
	parser: XMLParser;
	xmlBuilder: XMLBuilder;
	resolvePartPath: (slidePath: string, relationshipTarget: string) => string;
	updateNotesXmlText: (
		notesXmlObject: XmlObject,
		notesText: string,
		notesSegments?: PptxSlide['notesSegments'],
	) => boolean;
	compatibilityReporter: {
		reportWarning: (warning: CompatibilityWarningInput) => void;
	};
}

export interface IPptxSlideNotesPartUpdater {
	updateNotesPart(init: PptxSlideNotesPartUpdaterInput): Promise<void>;
}

export class PptxSlideNotesPartUpdater implements IPptxSlideNotesPartUpdater {
	public async updateNotesPart(init: PptxSlideNotesPartUpdaterInput): Promise<void> {
		const hasNotes = init.slide.notes !== undefined || (init.slide.notesSegments?.length ?? 0) > 0;

		const notesRelationship = init.relationshipRegistry.findFirstByTypeOrTargetIncludes(
			init.slideNotesRelationshipType,
			'notesslide',
		);
		if (!notesRelationship) {
			if (hasNotes) {
				await this.createNotesPart(init);
			}
			return;
		}

		const notesTarget = String(notesRelationship['@_Target'] || '').trim();
		if (notesTarget.length === 0) {
			return;
		}

		let notesPath = init.resolvePartPath(init.slide.id, notesTarget);
		const notesXml = await init.zip.file(notesPath)?.async('string');
		if (!notesXml) {
			if (hasNotes) {
				this.reportMissingNotesPart(init, notesPath);
			}
			return;
		}

		notesPath = await this.claimNotesPart(init, notesRelationship, notesPath, notesXml);
		if (!hasNotes) {
			return;
		}
		const notesXmlObject = init.parser.parse(notesXml) as XmlObject;
		const didUpdate = init.updateNotesXmlText(
			notesXmlObject,
			init.slide.notes ?? '',
			init.slide.notesSegments,
		);
		if (!didUpdate) {
			this.reportSkippedNotesUpdate(init);
			return;
		}

		init.zip.file(notesPath, init.xmlBuilder.build(notesXmlObject));
	}

	/**
	 * A notes slide belongs to one slide. A slide copied from another one shares the source's
	 * notes part: give it its own copy (or take the part over when its slide is gone), with a
	 * back-relationship to it.
	 */
	private async claimNotesPart(
		init: PptxSlideNotesPartUpdaterInput,
		notesRelationship: XmlObject,
		notesPath: string,
		notesXml: string,
	): Promise<string> {
		const relsPathOf = (path: string) => path.replace(/([^/]+)$/, '_rels/$1.rels');
		const relsXml = await init.zip.file(relsPathOf(notesPath))?.async('string');
		const tree = relsXml
			? (init.parser.parse(relsXml) as { Relationships: XmlObject })
			: {
					Relationships: {
						'@_xmlns': 'http://schemas.openxmlformats.org/package/2006/relationships',
					},
				};
		const raw = tree.Relationships.Relationship;
		const relationships = (Array.isArray(raw) ? raw : raw ? [raw] : []) as XmlObject[];
		let back = relationships.find((rel) => /\/slide$/u.test(String(rel['@_Type'] ?? '')));
		const owner = back ? resolvePartTarget(notesPath, String(back['@_Target'] ?? '')) : void 0;
		if (owner === init.slide.id) {
			return notesPath;
		}
		let shared = owner !== void 0 && Boolean(init.zip.file(owner));
		if (owner === void 0) {
			for (const path of Object.keys(init.zip.files).filter((path2) =>
				/^ppt\/slides\/_rels\/slide\d+\.xml\.rels$/u.test(path2),
			)) {
				const slide = path.replace('_rels/', '').replace(/\.rels$/u, '');
				if (
					slide !== init.slide.id &&
					init.zip.file(slide) &&
					(await init.zip.file(path)!.async('string')).includes(
						`notesSlides/${notesPath.slice(notesPath.lastIndexOf('/') + 1)}"`,
					)
				) {
					shared = true;
				}
			}
		}
		if (shared) {
			const copy = this.nextNotesPartPath(init.zip);
			init.zip.file(copy, notesXml);
			init.relationshipRegistry.upsertRelationship(
				String(notesRelationship['@_Id']),
				String(notesRelationship['@_Type']),
				`../notesSlides/${copy.slice(copy.lastIndexOf('/') + 1)}`,
			);
			await this.addNotesContentType(init, copy);
			notesPath = copy;
		} else if (owner === void 0) {
			return notesPath;
		}
		if (!back) {
			const used = new Set(relationships.map((rel) => String(rel['@_Id'])));
			let index = 1;
			while (used.has(`rId${index}`)) {
				index += 1;
			}
			back = {
				'@_Id': `rId${index}`,
				'@_Type': String(notesRelationship['@_Type']).replace(/notesSlide$/u, 'slide'),
			};
			relationships.push(back);
			tree.Relationships.Relationship = relationships;
		}
		back['@_Target'] = relativePartTarget(notesPath, init.slide.id);
		init.zip.file(relsPathOf(notesPath), init.xmlBuilder.build(tree));
		return notesPath;
	}

	private async createNotesPart(init: PptxSlideNotesPartUpdaterInput): Promise<void> {
		const notesMasterPath = Object.keys(init.zip.files).find((path) =>
			/^ppt\/notesMasters\/notesMaster\d+\.xml$/u.test(path),
		);
		if (!notesMasterPath) {
			this.reportMissingNotesRelationship(init);
			return;
		}
		const notesPath = this.nextNotesPartPath(init.zip);
		const fileName = notesPath.slice(notesPath.lastIndexOf('/') + 1);
		const strict = init.slideNotesRelationshipType.startsWith('http://purl.oclc.org/');
		const relationshipBase = strict
			? 'http://purl.oclc.org/ooxml/officeDocument/relationships/'
			: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
		const relationshipId = init.relationshipRegistry.nextRelationshipId();
		init.relationshipRegistry.upsertRelationship(
			relationshipId,
			init.slideNotesRelationshipType,
			`../notesSlides/${fileName}`,
		);
		init.zip.file(notesPath, init.xmlBuilder.build(buildNotesSlide(init.slide, strict)));
		init.zip.file(
			`ppt/notesSlides/_rels/${fileName}.rels`,
			init.xmlBuilder.build({
				Relationships: {
					'@_xmlns': 'http://schemas.openxmlformats.org/package/2006/relationships',
					Relationship: [
						{
							'@_Id': 'rId1',
							'@_Type': `${relationshipBase}notesMaster`,
							'@_Target': `../notesMasters/${notesMasterPath.slice(notesMasterPath.lastIndexOf('/') + 1)}`,
						},
						{
							'@_Id': 'rId2',
							'@_Type': `${relationshipBase}slide`,
							'@_Target': `../slides/${init.slide.id.slice(init.slide.id.lastIndexOf('/') + 1)}`,
						},
					],
				},
			}),
		);
		await this.addNotesContentType(init, notesPath);
	}

	private nextNotesPartPath(zip: JSZip): string {
		const used = new Set<number>();
		for (const path of Object.keys(zip.files)) {
			const index = /^ppt\/notesSlides\/notesSlide(?<index>\d+)\.xml$/u.exec(path)?.groups?.index;
			if (index) {
				used.add(Number.parseInt(index, 10));
			}
		}
		let index = 1;
		while (used.has(index)) {
			index += 1;
		}
		return `ppt/notesSlides/notesSlide${index}.xml`;
	}

	private async addNotesContentType(
		init: PptxSlideNotesPartUpdaterInput,
		notesPath: string,
	): Promise<void> {
		const xml = await init.zip.file('[Content_Types].xml')?.async('string');
		if (!xml) {
			return;
		}
		const data = init.parser.parse(xml) as XmlObject;
		const root = (data['Types'] ?? {}) as XmlObject;
		const raw = root['Override'];
		const overrides = Array.isArray(raw) ? (raw as XmlObject[]) : raw ? [raw as XmlObject] : [];
		overrides.push({
			'@_PartName': `/${notesPath}`,
			'@_ContentType':
				'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml',
		});
		root['Override'] = overrides;
		data['Types'] = root;
		init.zip.file('[Content_Types].xml', init.xmlBuilder.build(data));
	}

	private reportMissingNotesRelationship(init: PptxSlideNotesPartUpdaterInput): void {
		init.compatibilityReporter.reportWarning({
			code: 'SAVE_NOTES_RELATIONSHIP_MISSING',
			message:
				'Slide notes were edited, but the slide has no notes relationship. Notes update was skipped.',
			scope: 'save',
			slideId: init.slide.id,
		});
	}

	private reportMissingNotesPart(init: PptxSlideNotesPartUpdaterInput, notesPath: string): void {
		init.compatibilityReporter.reportWarning({
			code: 'SAVE_NOTES_PART_MISSING',
			message:
				'Speaker notes relationship exists but the notes part is missing. Notes update was skipped.',
			scope: 'save',
			slideId: init.slide.id,
			xmlPath: notesPath,
		});
	}

	private reportSkippedNotesUpdate(init: PptxSlideNotesPartUpdaterInput): void {
		init.compatibilityReporter.reportWarning({
			code: 'SAVE_NOTES_UPDATE_SKIPPED',
			message:
				'Speaker notes were present but no editable notes body was found. Notes were left unchanged.',
			scope: 'save',
			slideId: init.slide.id,
		});
	}
}
