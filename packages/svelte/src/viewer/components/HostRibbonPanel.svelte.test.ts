import { flushSync, mount, unmount } from 'svelte';
import { expect, it, vi } from 'vitest';

import { EditorState } from '../editor/editor-state.svelte';
import type { ViewerStateBag } from '../state/create-viewer-state-types';
import HostRibbonPanel from './HostRibbonPanel.svelte';

it('changes the slide background through the host panel and supports undo', async () => {
	const editor = new EditorState({ getCurrent: () => 0, getHandler: () => null });
	editor.editable = true;
	editor.setSlides([{ id: 's1', rId: 'r1', slideNumber: 1, elements: [] }]);
	const target = document.createElement('div');
	document.body.appendChild(target);
	const instance = mount(HostRibbonPanel, {
		target,
		props: { panel: 'background', vm: { editor } as ViewerStateBag, onclose: vi.fn() },
	});
	try {
		flushSync();
		const input = target.querySelector<HTMLInputElement>('input[type=color]')!;
		input.value = '#ff0000';
		input.dispatchEvent(new Event('input', { bubbles: true }));
		flushSync();
		expect(editor.slides[0].backgroundColor).toBe('#ff0000');
		editor.undo();
		expect(editor.slides[0].backgroundColor).toBeUndefined();
	} finally {
		await unmount(instance);
		target.remove();
	}
});
