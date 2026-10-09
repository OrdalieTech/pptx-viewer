import { orderedYMaps, reorderYArray } from './collaboration-order';
import type { YArrayLike, YjsFactories, YMapLike } from './collaboration-schema';

/** Keep one existing row when concurrent deletions remove every visible row. */
export function tableRows(array: YArrayLike): YMapLike[] {
	const rows = orderedYMaps(array);
	const visible = rows.filter((row) => !row.get('_deleted'));
	return visible.length ? visible : rows.slice(0, 1);
}

/** Missing IDs mean new objects, never an instruction to overwrite the same position. */
export function reconcileItems<T extends { collaborationId?: string }>(
	items: T[],
	array: YArrayLike,
	factories: YjsFactories,
	update: (item: T, map: YMapLike, index: number) => void,
	retainDeleted = false,
): void {
	const ids = items.map((item) => item.collaborationId ?? crypto.randomUUID());
	if (new Set(ids).size !== ids.length || ids.some((id) => !id))
		throw new Error('Duplicate or empty PPTX table collaboration ID');
	const wanted = new Set(ids);
	const existing = new Map(orderedYMaps(array).map((map) => [String(map.get('id')), map]));
	for (let i = array.length - 1; i >= 0; i--) {
		const map = array.get(i) as YMapLike;
		const id = String(map.get('id'));
		if (existing.get(id) !== map) array.delete(i, 1);
		else if (!wanted.has(id)) {
			if (retainDeleted) {
				if (!map.get('_deleted')) map.set('_deleted', true);
			} else array.delete(i, 1);
		}
	}
	items.forEach((item, index) => {
		let map = existing.get(ids[index]);
		if (!map) {
			map = factories.createMap();
			map.set('id', ids[index]);
			array.push([map]);
		}
		if (retainDeleted && map.get('_deleted')) map.delete('_deleted');
		update(item, map, index);
	});
	reorderYArray(array, [
		...ids,
		...orderedYMaps(array)
			.map((map) => String(map.get('id')))
			.filter((id) => !wanted.has(id)),
	]);
}

/** Imported frame sizes stay authored until the row structure is edited. */
export function readTableHeight(elementMap: YMapLike): number | undefined {
	const table = elementMap.get('tableData') as YMapLike | undefined;
	if (table?.get('_heightFromRows')) {
		const heights = tableRows(table.get('rows') as YArrayLike).map(
			(row) => JSON.parse(String(row.get('_data') ?? '{}')).height as number,
		);
		if (heights.length && heights.every((height) => Number.isFinite(height) && height > 0)) {
			return heights.reduce((sum, height) => sum + height, 0);
		}
	}
	return elementMap.get('height') as number | undefined;
}
