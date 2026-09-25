import { describe, expect, it, vi } from 'vitest';
import { createHostRibbonApi } from './host-ribbon-api';
import type { ViewerStateBag } from '../state/create-viewer-state-types';

describe('host ribbon commands', () => {
  function setup(editable = true) {
    const vm = {
      editingActive: editable, collab: { readOnly: false },
      editor: { inkOps: { setTool: vi.fn() }, insertElement: vi.fn(),
        transitionOps: { applyTransition: vi.fn() }, animationOps: { addAnimation: vi.fn() } },
      viewer: { current: 0, zoomToFit: vi.fn() },
      loader: { canvasSize: { width: 960, height: 540 } },
      parityUi: { preferences: { showRulers: false }, slideSorterOpen: false },
      chromeUi: { inspectorOpen: false, inspectorTab: 'properties' },
      onNotesToggle: vi.fn(), enterPresenterView: vi.fn(),
    };
    return { vm, run: createHostRibbonApi(vm as unknown as ViewerStateBag, vi.fn()) };
  }
  it('routes edits to native history-aware controllers', () => {
    const { vm, run } = setup();
    run('chart'); run('transition-fade'); run('animation-fade'); run('pen');
    expect(vm.editor.insertElement).toHaveBeenCalledWith(expect.objectContaining({ type: 'chart' }));
    expect(vm.editor.transitionOps.applyTransition).toHaveBeenCalledWith('fade', 700, false);
    expect(vm.editor.animationOps.addAnimation).toHaveBeenCalledWith('entrance', 'fadeIn');
    expect(vm.editor.inkOps.setTool).toHaveBeenCalledWith('pen');
  });
  it.each(['wipe', 'split', 'reveal', 'cut', 'cover', 'uncover'] as const)('applies %s through the native controller', (type) => {
    const { vm, run } = setup();
    run(`transition-${type}`);
    expect(vm.editor.transitionOps.applyTransition).toHaveBeenCalledWith(type, 700, false);
  });
  it('allows view commands but prevents mutations in read-only mode', () => {
    const { vm, run } = setup(false);
    run('chart'); run('pen'); run('transition-fade'); run('rulers'); run('sorter'); run('presenter');
    expect(vm.editor.insertElement).not.toHaveBeenCalled();
    expect(vm.editor.inkOps.setTool).not.toHaveBeenCalled();
    expect(vm.editor.transitionOps.applyTransition).not.toHaveBeenCalled();
    expect(vm.parityUi.preferences.showRulers).toBe(true);
    expect(vm.parityUi.slideSorterOpen).toBe(true);
    expect(vm.enterPresenterView).toHaveBeenCalledOnce();
    run('rulers');
    expect(vm.parityUi.preferences.showRulers).toBe(false);
  });
});
