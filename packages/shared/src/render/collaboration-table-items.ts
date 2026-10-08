import { orderedYMaps, reorderYArray } from './collaboration-order';
import type { YArrayLike, YjsFactories, YMapLike } from './collaboration-schema';

/** Missing IDs mean new objects, never an instruction to overwrite the same position. */
export function reconcileItems<T extends { collaborationId?: string }>(
	items: T[],
	array: YArrayLike,
	factories: YjsFactories,
	update: (item: T, map: YMapLike, index: number) => void,
): void {
	const ids = items.map((item) => item.collaborationId ?? crypto.randomUUID());
	if (new Set(ids).size !== ids.length || ids.some((id) => !id))
		throw new Error('Duplicate or empty PPTX table collaboration ID');
	const wanted = new Set(ids);
	const existing = new Map(orderedYMaps(array).map((map) => [String(map.get('id')), map]));
	for (let i = array.length - 1; i >= 0; i--) {
		const map = array.get(i) as YMapLike;
		const id = String(map.get('id'));
		if (!wanted.has(id) || existing.get(id) !== map) array.delete(i, 1);
	}
	items.forEach((item, index) => {
		let map = existing.get(ids[index]);
		if (!map) {
			map = factories.createMap();
			map.set('id', ids[index]);
			array.push([map]);
		}
		update(item, map, index);
	});
	reorderYArray(array, ids);
}
