import type { PptxTableData } from 'pptx-viewer-core';

import { orderedYMaps } from './collaboration-order';
import type { YArrayLike, YMapLike, YjsFactories } from './collaboration-schema';

interface Column {
	id: string;
	width: number;
	order: number;
}
const PREFIX = '_column:';

/** Resolve the visible cell, creating only a missing concurrent row/column intersection. */
export function resolveTableCell(
	element: YMapLike,
	cellId: string,
	factories: YjsFactories,
): YMapLike | undefined {
	const table = element.get('tableData') as YMapLike | undefined;
	if (!table) return undefined;
	const columns = tableColumns(table);
	for (const row of orderedYMaps(table.get('rows') as YArrayLike)) {
		const cells = row.get('cells') as YArrayLike;
		// Concurrent first writes can create the same logical cell. Match the renderer's winner.
		const existing = orderedYMaps(cells)
			.reverse()
			.find((cell) => cell.get('id') === cellId);
		if (existing) return existing;
		const column = columns.find((column) => `${row.get('id')}:${column.id}` === cellId);
		if (!column) continue;
		const cell = factories.createMap();
		cell.set('id', cellId);
		cell.set('_column', column.id);
		cell.set('_data', '{}');
		cell.set('textBody', factories.createText());
		cells.push([cell]);
		return cell;
	}
	return undefined;
}

/** Primitive entries also let two clients upgrade an existing table without replacing a shared type. */
export function tableColumns(table: YMapLike): Column[] {
	const columns: Column[] = [];
	table.forEach((value, key) => {
		if (
			key.startsWith(PREFIX) &&
			typeof value === 'string' &&
			!table.get(`_deletedColumn:${key.slice(PREFIX.length)}`)
		) {
			columns.push({ id: key.slice(PREFIX.length), ...JSON.parse(value) });
		}
	});
	return columns.sort((a, b) => a.order - b.order || a.id.localeCompare(b.id));
}

export function reconcileTableColumns(
	data: PptxTableData,
	table: YMapLike,
	elementId: string,
): string[] {
	let columns = tableColumns(table);
	const rows = table.get('rows') as YArrayLike;
	const oldWidths = (JSON.parse(String(table.get('_data'))) as PptxTableData).columnWidths;
	if (!columns.length) {
		const total = oldWidths.reduce((sum, width) => sum + width, 0);
		columns = oldWidths.map((width, order) => ({
			id: `${elementId}:column:${order}`,
			width: width / total,
			order,
		}));
		for (const column of columns) {
			table.set(PREFIX + column.id, JSON.stringify({ width: column.width, order: column.order }));
		}
	}
	const cellColumns = new Map<string, string>();
	for (const row of orderedYMaps(rows)) {
		orderedYMaps(row.get('cells') as YArrayLike).forEach((cell, index) => {
			const id = cell.get('_column') ?? columns[index]?.id;
			if (typeof id === 'string') {
				if (cell.get('_column') !== id) {
					cell.set('_column', id);
				}
				cellColumns.set(String(cell.get('id')), id);
			}
		});
	}
	const ids = data.columnWidths.map((_, index) => {
		const retained = data.rows
			.map((row) => cellColumns.get(row.cells[index]?.collaborationId ?? ''))
			.filter((id) => id !== undefined);
		return retained[0] ?? crypto.randomUUID();
	});
	if (new Set(ids).size !== ids.length) {
		throw new Error('PPTX table columns must have distinct identities');
	}
	const wanted = new Set(ids);
	for (const column of columns) {
		if (!wanted.has(column.id)) table.set(`_deletedColumn:${column.id}`, true);
	}
	const total = data.columnWidths.reduce((sum, width) => sum + width, 0);
	let previous = -Infinity;
	ids.forEach((id, index) => {
		const existing = columns.find((column) => column.id === id);
		const following = ids
			.slice(index + 1)
			.map((next) => columns.find((column) => column.id === next)?.order)
			.find((order) => order !== undefined && order > previous);
		const order =
			existing && existing.order > previous
				? existing.order
				: previous === -Infinity
					? (following ?? 0) - 1
					: following === undefined
						? previous + 1
						: (previous + following) / 2;
		const value = JSON.stringify({ width: data.columnWidths[index] / total, order });
		if (table.get(PREFIX + id) !== value) {
			table.set(PREFIX + id, value);
		}
		previous = order;
	});
	return ids;
}

/** Read every row in the same column order, including concurrent insertions. */
export function alignTableColumns(data: PptxTableData, table: YMapLike): PptxTableData {
	const columns = tableColumns(table);
	if (!columns.length) {
		return data;
	}
	const rows = orderedYMaps(table.get('rows') as YArrayLike);
	const total = columns.reduce((sum, column) => sum + column.width, 0);
	const authoredTotal = data.columnWidths.reduce((sum, width) => sum + width, 0);
	return {
		...data,
		columnWidths: columns.map((column) => (column.width * authoredTotal) / total),
		rows: data.rows.map((row, rowIndex) => {
			const maps = orderedYMaps(rows[rowIndex].get('cells') as YArrayLike);
			const cells = new Map(maps.map((cell, index) => [cell.get('_column'), row.cells[index]]));
			return {
				...row,
				cells: columns.map(
					(column) =>
						cells.get(column.id) ?? {
							text: '',
							collaborationId: `${row.collaborationId}:${column.id}`,
						},
				),
			};
		}),
	};
}
