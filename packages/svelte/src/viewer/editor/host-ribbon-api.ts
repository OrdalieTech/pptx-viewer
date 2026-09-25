import type { ViewerStateBag } from '../state/create-viewer-state-types';
import { playSlideTransitionPreview } from 'pptx-viewer-shared';
import { buildChartInsertElement } from './editor-insert-chart';
import { previewElementAnimation } from '../components/ribbon/animations/animation-preview-player';

/** Commands for host-owned ribbons; all edits use the native controllers. */
export type HostRibbonCommand = 'select' | 'pen' | 'highlighter' | 'eraser' | 'chart'
  | 'transition-none' | 'transition-morph' | 'transition-fade' | 'transition-push' | 'transition-all'
  | 'transition-wipe' | 'transition-split' | 'transition-reveal' | 'transition-cut' | 'transition-cover' | 'transition-uncover'
  | 'animation-appear' | 'animation-fade' | 'properties' | 'background'
  | 'comments' | 'notes' | 'rulers' | 'sorter' | 'presenter' | 'fit'
  | 'themes' | 'layout' | 'accessibility' | 'animations' | 'spelling' | 'animation-preview' | 'transition-preview';

export function createHostRibbonApi(vm: ViewerStateBag, openPanel: (panel: string) => void, getRoot: () => ParentNode | undefined = () => undefined) {
  return (command: HostRibbonCommand): void => {
    const { editor, viewer, parityUi, chromeUi } = vm;
    switch (command) {
      case 'transition-preview': {
        playSlideTransitionPreview(editor.slides[viewer.current]?.transition, document);
        return;
      }
      case 'animation-preview': {
        const animation = editor.slides[viewer.current]?.animations?.find((item) => item.elementId === editor.selectedElementId);
        const root = getRoot();
        if (animation && root) previewElementAnimation(animation, root);
        return;
      }
      case 'notes': vm.onNotesToggle(); return;
      case 'rulers': parityUi.preferences = { ...parityUi.preferences, showRulers: !parityUi.preferences.showRulers }; return;
      case 'sorter': parityUi.slideSorterOpen = true; return;
      case 'presenter': vm.enterPresenterView(); return;
      case 'fit': viewer.zoomToFit(); return;
      case 'comments': chromeUi.inspectorOpen = true; chromeUi.inspectorTab = 'comments'; return;
      case 'background': case 'themes': case 'layout': case 'accessibility': case 'animations': openPanel(command); return;
      case 'spelling': parityUi.preferences = { ...parityUi.preferences, spellCheck: !parityUi.preferences.spellCheck }; return;
      case 'properties': chromeUi.inspectorOpen = true; chromeUi.inspectorTab = 'properties'; return;
    }
    if (!vm.editingActive || vm.collab.readOnly) return;
    switch (command) {
      case 'select': case 'pen': case 'highlighter': case 'eraser': editor.inkOps.setTool(command); return;
      case 'chart': editor.insertElement(buildChartInsertElement('column', vm.loader.canvasSize)); return;
      case 'animation-appear': editor.animationOps.addAnimation('entrance', 'appear'); return;
      case 'animation-fade': editor.animationOps.addAnimation('entrance', 'fadeIn'); return;
      case 'transition-none': case 'transition-morph': case 'transition-fade': case 'transition-push':
      case 'transition-wipe': case 'transition-split': case 'transition-reveal': case 'transition-cut': case 'transition-cover': case 'transition-uncover':
        editor.transitionOps.applyTransition(command.slice('transition-'.length) as import('pptx-viewer-core').PptxTransitionType, 700, false); return;
      case 'transition-all': {
        const transition = editor.slides[viewer.current]?.transition;
        if (transition) editor.transitionOps.applyTransition(transition.type, transition.durationMs ?? 700, true);
      }
    }
  };
}
