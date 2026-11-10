import { describe, expect, it } from 'vitest';

import type { XmlObject } from '../../types';
import {
	annotateSmartArtTextOrder,
	orderedSmartArtTextEntries,
	smartArtChildOrder,
} from './smartart-text-order';

describe('smartArt text order annotation', () => {
	it.each([null, undefined, '', 0, false])('ignores primitive parser result %j', (parsed) => {
		expect(() => annotateSmartArtTextOrder('<a:txBody><a:p/></a:txBody>', parsed)).not.toThrow();
	});

	it('filters primitive paragraph array values before WeakMap annotation', () => {
		const paragraph: XmlObject = {
			'a:r': [{ 'a:t': 'First' }, { 'a:t': 'Second' }],
			'a:tab': '',
		};
		const parsed = {
			'a:txBody': {
				'a:p': [null, '', 7, paragraph],
			},
		};

		expect(() =>
			annotateSmartArtTextOrder(
				'<a:txBody><a:p><a:r><a:t>First</a:t></a:r><a:tab/><a:r><a:t>Second</a:t></a:r></a:p></a:txBody>',
				parsed,
			),
		).not.toThrow();
		expect(smartArtChildOrder(paragraph)).toStrictEqual(['r', 'tab', 'r']);
		expect(orderedSmartArtTextEntries(paragraph).map(([key]) => key)).toStrictEqual([
			'a:r',
			'a:tab',
			'a:r',
		]);
	});
	// The walk that replaced the single global regex must find the same spans,
	// false matches included (`<p:sp>` reads as an element `p`, `<a:p/>` runs to
	// the next `</a:p>`): the annotation feeds the saved XML.
	describe('matches the reference regex', () => {
		const referenceChildOrder = (xml: string): string[] => {
			const order: string[] = [];
			let depth = 0;
			for (const match of xml.matchAll(/<(\/)?([A-Za-z_][\w.-]*:)?([A-Za-z][\w.-]*)\b[^>]*>/gu)) {
				if (match[1]) {
					depth = Math.max(0, depth - 1);
					continue;
				}
				if (depth === 0) {
					order.push(match[3]);
				}
				if (!/\/\s*>$/u.test(match[0])) {
					depth++;
				}
			}
			return order;
		};
		const referenceOrders = (xml: string, name: string): string[][] =>
			[
				...xml.matchAll(
					new RegExp(`<([A-Za-z_][\\w.-]*:)?${name}\\b[^>]*>([\\s\\S]*?)<\\/\\1${name}\\s*>`, 'gu'),
				),
			].map((match) => referenceChildOrder(match[2]));
		const run = '<a:r><a:rPr lang="en-US"/><a:t>Text</a:t></a:r>';
		const shape = (body: string): string =>
			`<p:sp><p:nvSpPr><p:cNvPr id="2" name="a > b"/></p:nvSpPr><p:txBody><a:bodyPr/>${body}</p:txBody></p:sp>`;

		it.each([
			[
				'a slide',
				`<p:sld><p:cSld><p:spTree>${shape(`<a:p><a:pPr lvl="1"/>${run}<a:br/>${run}<a:fld id="{1}" type="slidenum"><a:t>1</a:t></a:fld><a:endParaRPr/></a:p>`).repeat(3)}</p:spTree></p:cSld></p:sld>`,
			],
			[
				'self-closing paragraphs and breaks',
				shape(`<a:p/><a:p>${run}<a:br/></a:p><a:p/><a:br><a:rPr/></a:br><a:p>${run}</a:p>`),
			],
			[
				'unprefixed elements',
				`<txBody><p><r><t>One</t></r><br/><r><t>Two</t></r></p ><p:sp/><p><fld/></p></txBody>`,
			],
			[
				'a prefix that never closes',
				`<p:txBody>${'<p:sp><r:id/></p:sp>'.repeat(5)}<a:p>${run}</a:p></p:txBody>`,
			],
			['closing tags without a start', `</a:p><a:txBody></a:r><a:p>${run}</a:p></a:br></a:txBody>`],
			[
				'lookalike names',
				`<a:txBody><a:pPr/><a:p-x></a:p-x><x.y:p><a:r/></x.y:p><a:rPr></a:rPr><a:p>${run}</a:p></a:txBody>`,
			],
			['nested paragraphs', `<a:txBody><a:p><b:p>${run}</b:p><a:p>${run}</a:p></a:p></a:txBody>`],
		])('%s', (_name, xml) => {
			const nodes = (count: number): XmlObject[] => Array.from({ length: count }, () => ({}));
			const first: XmlObject = { 'a:r': nodes(40), 'a:fld': nodes(40), 'a:br': nodes(40) };
			const paragraphs = [first, ...nodes(40)];
			annotateSmartArtTextOrder(xml, { 'a:txBody': { 'a:p': paragraphs } });

			const found = (items: XmlObject[]) => items.map(smartArtChildOrder).filter(Boolean);
			expect(found(paragraphs)).toStrictEqual(referenceOrders(xml, 'p'));
			for (const name of ['r', 'fld', 'br']) {
				expect(found(first[`a:${name}`] as XmlObject[])).toStrictEqual(referenceOrders(xml, name));
			}
		});
	});
});
