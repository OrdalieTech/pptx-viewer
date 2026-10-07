import type { PptxSlide, XmlObject } from '../../types';

export function buildNotesSlide(slide: PptxSlide, strict: boolean): XmlObject {
	const notesText =
		slide.notes ?? slide.notesSegments?.map((segment) => String(segment.text ?? '')).join('') ?? '';
	const p = strict
		? 'http://purl.oclc.org/ooxml/presentationml/main'
		: 'http://schemas.openxmlformats.org/presentationml/2006/main';
	const a = strict
		? 'http://purl.oclc.org/ooxml/drawingml/main'
		: 'http://schemas.openxmlformats.org/drawingml/2006/main';
	const r = strict
		? 'http://purl.oclc.org/ooxml/officeDocument/relationships'
		: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
	return {
		'p:notes': {
			'@_xmlns:a': a,
			'@_xmlns:r': r,
			'@_xmlns:p': p,
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
					'p:sp': {
						'p:nvSpPr': {
							'p:cNvPr': { '@_id': '2', '@_name': 'Notes Placeholder' },
							'p:cNvSpPr': {},
							'p:nvPr': { 'p:ph': { '@_type': 'body', '@_idx': '1' } },
						},
						'p:spPr': {},
						'p:txBody': {
							'a:bodyPr': {},
							'a:lstStyle': {},
							'a:p': {
								'a:r': { 'a:rPr': { '@_lang': 'en-US' }, 'a:t': notesText },
								'a:endParaRPr': { '@_lang': 'en-US' },
							},
						},
					},
				},
			},
			'p:clrMapOvr': { 'a:masterClrMapping': {} },
		},
	};
}
