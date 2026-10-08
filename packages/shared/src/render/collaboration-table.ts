import type {
	PptxTableCell,
	PptxTableCellTextRun,
	PptxTableData,
	PptxTableRow,
	TextSegment,
} from 'pptx-viewer-core';

import { orderedYMaps } from './collaboration-order';
import type { YArrayLike, YjsFactories, YMapLike } from './collaboration-sync';
import { alignTableColumns, reconcileTableColumns } from './collaboration-table-columns';
import { reconcileItems } from './collaboration-table-items';
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

function tableSegments(cell: PptxTableCell): TextSegment[] {
	return cell.textSegments ?? [{ text: cell.text, style: {} }];
}

function tableRuns(segments: TextSegment[]): PptxTableCellTextRun[] {
	return segments.map((segment) => ({
		text: segment.text,
		isParagraphBreak: segment.isParagraphBreak,
		isLineBreak: segment.isLineBreak,
		bold: segment.style.bold,
		italic: segment.style.italic,
		underline: segment.style.underline,
		strikethrough: segment.style.strikethrough,
		color: segment.style.color,
		fontSize:
			typeof segment.style.fontSize === 'number' ? segment.style.fontSize * 0.75 : undefined,
		fontFamily: segment.style.fontFamily,
	}));
}

function isLegacyTableCell(
	cell: RecordValue,
	text: { toString: () => string },
	segments: Record<string, unknown>[],
): boolean {
	const runs = cell.textRuns as PptxTableCellTextRun[] | undefined;
	return (
		Array.isArray(runs) &&
		runs.length > 0 &&
		segments.length === 1 &&
		Object.keys((segments[0].style as RecordValue | undefined) ?? {}).length === 0 &&
		!segments[0].paragraphProperties &&
		runs
			.map((run) => (run.isParagraphBreak || run.isLineBreak ? '\n' : String(run.text ?? '')))
			.join('') === text.toString()
	);
}

function reconcileTableText(
	text: import('./collaboration-text-merge').YTextEditableLike,
	cell: PptxTableCell,
): void {
	const segments = tableSegments(cell);
	const rendered = segments
		.map((segment) => (segment.isParagraphBreak || segment.isLineBreak ? '\n' : segment.text))
		.join('');
	if (rendered === cell.text) {
		mergeDeltaIntoYText(text, encodeSegmentsToDelta(segments));
		return;
	}

	// Template editing changes the flat text but can leave its old segment
	// model attached. Apply only the changed span so the surrounding run styles
	// and paragraph properties stay on the live Y.Text.
	const before = text.toString();
	const after = cell.text;
	let start = 0;
	let end = before.length;
	let nextEnd = after.length;
	while (start < end && start < nextEnd && before[start] === after[start]) start++;
	while (end > start && nextEnd > start && before[end - 1] === after[nextEnd - 1]) {
		end--;
		nextEnd--;
	}
	if (start === end && start === nextEnd) return;

	let offset = 0;
	let attributes: Record<string, string> = {};
	for (const run of text.toDelta()) {
		const length = typeof run.insert === 'string' ? run.insert.length : 1;
		if (start < offset + length || start === before.length) {
			attributes = (run.attributes ?? {}) as Record<string, string>;
		}
		if (start < offset + length) break;
		offset += length;
	}
	const inherited = { ...attributes };
	delete inherited.pb;
	delete inherited.lb;
	const replacement = after.slice(start, nextEnd).split('\n');
	if (end > start) text.delete(start, end - start);
	for (let index = 0; index < replacement.length; index++) {
		if (index > 0) {
			text.insert(start, '\n', { ...inherited, pb: '1' });
			start++;
		}
		if (replacement[index]) {
			text.insert(start, replacement[index], inherited);
			start += replacement[index].length;
		}
	}
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
			cellMap.set(
				'_data',
				metadata(cell as unknown as RecordValue, [
					'text',
					'textSegments',
					'textRuns',
					'collaborationId',
				]),
			);
			const text = factories.createText();
			encodeTextBody(tableSegments(cell), text);
			cellMap.set('textBody', text);
			cells.push([cellMap]);
		});
		rowMap.set('cells', cells);
		rows.push([rowMap]);
	});
	table.set('rows', rows);
	elementMap.set('tableData', table);
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
	const columns = reconcileTableColumns(data, table, elementId);
	setMetadata(table, metadata(data as unknown as RecordValue, ['rows']));
	const rows = table.get('rows') as YArrayLike;
	const rowIds = new Set(orderedYMaps(rows).map((row) => String(row.get('id'))));
	if (
		data.rows.length !== rows.length ||
		data.rows.some((row) => !rowIds.has(row.collaborationId ?? ''))
	) {
		table.set('_heightFromRows', true);
	}
	reconcileItems(data.rows, rows, factories, (row, rowMap) => {
		if (!rowMap.get('cells')) rowMap.set('cells', factories.createArray());
		setMetadata(rowMap, metadata(row as unknown as RecordValue, ['cells', 'collaborationId']));
		const cells = rowMap.get('cells') as YArrayLike;
		reconcileItems(row.cells, cells, factories, (cell, cellMap, index) => {
			if (cellMap.get('_column') !== columns[index]) cellMap.set('_column', columns[index]);
			const text = cellMap.get('textBody');
			const previous = JSON.parse(String(cellMap.get('_data') ?? '{}')) as RecordValue;
			if (
				!Array.isArray(cell.textSegments) &&
				Array.isArray(previous.textRuns) &&
				JSON.stringify(cell.textRuns) === JSON.stringify(previous.textRuns) &&
				isYTextEditable(text) &&
				cell.text === text.toString()
			) {
				const retained = JSON.parse(
					metadata(cell as unknown as RecordValue, [
						'text',
						'textSegments',
						'textRuns',
						'collaborationId',
					]),
				) as RecordValue;
				retained.textRuns = previous.textRuns;
				setMetadata(cellMap, JSON.stringify(retained));
				return;
			}
			setMetadata(
				cellMap,
				metadata(cell as unknown as RecordValue, [
					'text',
					'textSegments',
					'textRuns',
					'collaborationId',
				]),
			);
			const segments = tableSegments(cell);
			if (isYTextEditable(text)) {
				reconcileTableText(text, cell);
			} else {
				const next = factories.createText();
				encodeTextBody(segments, next);
				cellMap.set('textBody', next);
			}
		});
	});
}

/** Imported frame sizes stay authored until the row structure is edited. */
export function readTableHeight(elementMap: YMapLike): number | undefined {
	const table = elementMap.get('tableData') as YMapLike | undefined;
	if (table?.get('_heightFromRows')) {
		const heights = orderedYMaps(table.get('rows') as YArrayLike).map(
			(row) => JSON.parse(String(row.get('_data') ?? '{}')).height as number,
		);
		if (heights.length && heights.every((height) => Number.isFinite(height) && height > 0)) {
			return heights.reduce((sum, height) => sum + height, 0);
		}
	}
	return elementMap.get('height') as number | undefined;
}

export function readTableData(elementMap: YMapLike): PptxTableData | undefined {
	const table = elementMap.get('tableData') as YMapLike | undefined;
	if (!table) return undefined;
	const readMetadata = <T extends object>(map: YMapLike): T =>
		JSON.parse(map.get('_data') as string);
	const rows = table.get('rows') as YArrayLike;
	return alignTableColumns(
		{
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
						const cellMetadata = readMetadata<Omit<PptxTableCell, 'text'>>(cell);
						const legacy = isLegacyTableCell(cellMetadata as RecordValue, text, segments);
						return {
							...cellMetadata,
							collaborationId: String(cell.get('id')),
							...(!legacy ? { textSegments: segments as unknown as TextSegment[] } : {}),
							textRuns: legacy
								? cellMetadata.textRuns
								: tableRuns(segments as unknown as TextSegment[]),
							text: segments
								.map((segment) =>
									segment.isParagraphBreak || segment.isLineBreak ? '\n' : segment.text,
								)
								.join(''),
						};
					}),
				};
			}),
		},
		table,
	);
}
