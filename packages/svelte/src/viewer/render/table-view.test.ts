import { expect, test } from 'vitest';

import { buildTableRows } from './table-view';

test('table typography uses deck spacing and font fallbacks', () => {
	const rows = buildTableRows({
		columnWidths: [1],
		rows: [
			{
				cells: [
					{
						text: 'Row',
						style: { fontFamily: 'Segoe UI Light' },
						textSegments: [{ text: 'Row', paragraphProperties: { lineSpacing: 1.15 } }],
					},
				],
			},
		],
	});
	expect(rows[0].cells[0].style).toContain('line-height: 1.15');
	expect(rows[0].cells[0].style).toContain('sans-serif');
});
