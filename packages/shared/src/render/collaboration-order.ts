import type { YArrayLike, YMapLike } from './collaboration-sync';

function rank(map: YMapLike, fallback: number): number {
	const value = map.get('_order');
	return typeof value === 'number' && Number.isFinite(value) ? value : fallback;
}

/** Ordering is separate from map identity, so moving a slide keeps concurrent text edits. */
export function orderedYMaps(array: YArrayLike): YMapLike[] {
	return array
		.toArray()
		.map((map, index) => ({ map: map as YMapLike, index }))
		.sort(
			(a, b) =>
				rank(a.map, a.index) - rank(b.map, b.index) ||
				String(a.map.get('id')).localeCompare(String(b.map.get('id'))),
		)
		.map(({ map }) => map);
}

export function reorderYArray(array: YArrayLike, ids: readonly string[]): void {
	let current = orderedYMaps(array);
	const maps = new Map(current.map((map) => [String(map.get('id')), map]));
	if (
		new Set(ids).size !== ids.length ||
		ids.length !== current.length ||
		ids.some((id) => !maps.has(id))
	) {
		throw new Error('PPTX reorder requires each current ID exactly once');
	}
	for (let i = 0; i < ids.length; i++) {
		if (current[i].get('id') === ids[i]) {
			continue;
		}
		const before = i ? rank(current[i - 1], i - 1) : undefined;
		const after = rank(current[i], i);
		const value = before === undefined ? after - 1 : (before + after) / 2;
		if (!Number.isFinite(value) || value === before || value === after) {
			// Extremely dense ranks: compact positions without rebuilding any shared map.
			ids.forEach((id, index) => maps.get(id)!.set('_order', index));
			return;
		}
		maps.get(ids[i])!.set('_order', value);
		current = orderedYMaps(array);
	}
}
