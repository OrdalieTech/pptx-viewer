import type { PptxTableData, PptxTableRow, PptxTableCell } from 'pptx-viewer-core';

import { orderedYMaps, reorderYArray } from './collaboration-order';
import type { YArrayLike, YjsFactories, YMapLike } from './collaboration-sync';
import {
	decodeTextBody,
	encodeSegmentsToDelta,
	encodeTextBody,
	isYTextLike,
} from './collaboration-text-codec';
import { isYTextEditable, mergeDeltaIntoYText } from './collaboration-text-merge';

type RecordValue = Record<string, unknown>;

function metadata(value: RecordValue, excluded: string[]): string {
	return JSON.stringify(
		Object.fromEntries(Object.entries(value).filter(([key]) => !excluded.includes(key))),
	);
}

function setMetadata(map: YMapLike, data: string): void {
	if (map.get('_data') !== data) map.set('_data', data);
}

function validateIds(data: PptxTableData): void {
	const ids = new Set<string>();
	for (const row of data.rows) {
		for (const item of [row, ...row.cells]) {
			const id = item.collaborationId;
			if (id === undefined) continue;
			if (!id || ids.has(id)) throw new Error('Duplicate or empty PPTX table collaboration ID');
			ids.add(id);
		}
	}
}

/** Tables retain structural metadata separately from collaboratively editable cell text. */
export function writeTableData(
	data: PptxTableData,
	elementMap: YMapLike,
	factories: YjsFactories,
	elementId: string,
): void {
	validateIds(data);
	const table = factories.createMap();
	table.set('_data', metadata(data as unknown as RecordValue, ['rows']));
	const rows = factories.createArray();
	data.rows.forEach((row, rowIndex) => {
		const rowMap = factories.createMap();
		rowMap.set('id', row.collaborationId ?? `${elementId}:row:${rowIndex}`);
		rowMap.set('_data', metadata(row as unknown as RecordValue, ['cells', 'collaborationId']));
		const cells = factories.createArray();
		row.cells.forEach((cell, columnIndex) => {
			const cellMap = factories.createMap();
			cellMap.set('id', cell.collaborationId ?? `${elementId}:cell:${rowIndex}:${columnIndex}`);
			cellMap.set('_data', metadata(cell as unknown as RecordValue, ['text', 'collaborationId']));
			const text = factories.createText();
			encodeTextBody([{ text: cell.text, style: {} }], text);
			cellMap.set('textBody', text);
			cells.push([cellMap]);
		});
		rowMap.set('cells', cells);
		rows.push([rowMap]);
	});
	table.set('rows', rows);
	elementMap.set('tableData', table);
}

/** Missing IDs mean new objects, never an instruction to overwrite the same position. */
function reconcileItems<T extends { collaborationId?: string }>(
	items: T[],
	array: YArrayLike,
	factories: YjsFactories,
	update: (item: T, map: YMapLike) => void,
): void {
	const ids = items.map((item) => item.collaborationId ?? crypto.randomUUID());
	if (new Set(ids).size !== ids.length || ids.some((id) => !id))
		throw new Error('Duplicate or empty PPTX table collaboration ID');
	const wanted = new Set(ids);
	const existing = new Map(orderedYMaps(array).map((map) => [String(map.get('id')), map]));
	for (let i = array.length - 1; i >= 0; i--)
		if (!wanted.has(String((array.get(i) as YMapLike).get('id')))) array.delete(i, 1);
	items.forEach((item, index) => {
		let map = existing.get(ids[index]);
		if (!map) {
			map = factories.createMap();
			map.set('id', ids[index]);
			array.push([map]);
		}
		update(item, map);
	});
	reorderYArray(array, ids);
}

export function reconcileTableData(
	data: PptxTableData,
	elementMap: YMapLike,
	factories: YjsFactories,
	elementId: string,
): void {
	validateIds(data);
	const table = elementMap.get('tableData') as YMapLike | undefined;
	if (!table) return writeTableData(data, elementMap, factories, elementId);
	setMetadata(table, metadata(data as unknown as RecordValue, ['rows']));
	const rows = table.get('rows') as YArrayLike;
	reconcileItems(data.rows, rows, factories, (row, rowMap) => {
		if (!rowMap.get('cells')) rowMap.set('cells', factories.createArray());
		setMetadata(rowMap, metadata(row as unknown as RecordValue, ['cells', 'collaborationId']));
		const cells = rowMap.get('cells') as YArrayLike;
		reconcileItems(row.cells, cells, factories, (cell, cellMap) => {
			setMetadata(cellMap, metadata(cell as unknown as RecordValue, ['text', 'collaborationId']));
			const segments = [{ text: cell.text, style: {} }];
			const text = cellMap.get('textBody');
			if (isYTextEditable(text)) {
				mergeDeltaIntoYText(text, encodeSegmentsToDelta(segments));
			} else {
				const next = factories.createText();
				encodeTextBody(segments, next);
				cellMap.set('textBody', next);
			}
		});
	});
}

export function readTableData(elementMap: YMapLike): PptxTableData | undefined {
	const table = elementMap.get('tableData') as YMapLike | undefined;
	if (!table) return undefined;
	const readMetadata = <T extends object>(map: YMapLike): T =>
		JSON.parse(map.get('_data') as string);
	const rows = table.get('rows') as YArrayLike;
	return {
		...readMetadata<Omit<PptxTableData, 'rows'>>(table),
		rows: orderedYMaps(rows).map((row) => {
			const cells = row.get('cells') as YArrayLike;
			return {
				...readMetadata<Omit<PptxTableRow, 'cells'>>(row),
				collaborationId: String(row.get('id')),
				cells: orderedYMaps(cells).map((cell) => {
					const text = cell.get('textBody');
					if (!isYTextLike(text)) throw new Error('Incompatible PPTX table cell text');
					const segments = decodeTextBody(text);
					return {
						...readMetadata<Omit<PptxTableCell, 'text'>>(cell),
						collaborationId: String(cell.get('id')),
						text: segments
							.map((segment) => (segment.isParagraphBreak ? '\n' : segment.text))
							.join(''),
					};
				}),
			};
		}),
	};
}
