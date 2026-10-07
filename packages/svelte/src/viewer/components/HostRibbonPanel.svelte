<script lang="ts">
  import type { ViewerStateBag } from '../state/create-viewer-state-types';
  import FormatBackgroundPanel from './ribbon/design/FormatBackgroundPanel.svelte';
  import ReviewAccessibilityPanel from './ribbon/review/ReviewAccessibilityPanel.svelte';
  import ThemeSection from './inspector/ThemeSection.svelte';
  import LayoutGalleryMenu from './ribbon/home/LayoutGalleryMenu.svelte';
  import AnimationPanel from './inspector/AnimationPanel.svelte';
  const { vm, panel, onclose }: { vm: ViewerStateBag; panel: string; onclose: () => void } = $props();
  const titles: Record<string, string> = { background: 'Arrière-plan', themes: 'Thèmes et variantes du document', layout: 'Mise en page', accessibility: 'Accessibilité', animations: 'Animations et minutage' };
  const layouts = panel === 'layout' ? Promise.all([vm.editor.slidesOps.availableLayouts(), vm.editor.slidesOps.layoutPreviews()]) : null;
  let failure = $state('');
  function show(node: HTMLDialogElement) { node.showModal(); }
</script>

<dialog class="host-panel" use:show aria-label={titles[panel]} oncancel={onclose} onkeydown={(event) => event.stopPropagation()}>
  <header><strong>{titles[panel]}</strong><button aria-label="Fermer le panneau" onclick={onclose}>×</button></header>
  {#if failure}<p role="alert">{failure}</p>{/if}
  {#if panel === 'background'}
    <FormatBackgroundPanel editor={vm.editor} open={true} {onclose} />
  {:else if panel === 'accessibility'}
    <ReviewAccessibilityPanel slides={vm.editor.slides} onnavigate={(index, id) => { vm.viewer.goTo(index); if (id) vm.editor.select(id); }} />
  {:else if panel === 'themes' && vm.loader.handler}
    <ThemeSection editor={vm.editor} handler={vm.loader.handler} theme={vm.loader.presentationTheme} expanded={true} onthemechange={(theme) => { vm.loader.presentationTheme = theme; vm.loader.colorScheme = theme.colorScheme; }} />
  {:else if panel === 'layout' && layouts}
    {#await layouts}<p>Chargement…</p>{:then [options, previews]}
      <LayoutGalleryMenu layouts={options} {previews} onselect={async (layout) => { try { await vm.editor.slidesOps.applyLayout(layout.path); onclose(); } catch (error) { failure = String(error); } }} />
    {:catch error}<p role="alert">{String(error)}</p>{/await}
  {:else if panel === 'animations'}
    <AnimationPanel editor={vm.editor} />
  {/if}
</dialog>

<style>
  .host-panel { position: fixed; margin: auto; width: min(480px, calc(100vw - 32px)); max-height: calc(100dvh - 64px); overflow: auto; padding: 12px; background: var(--pptx-card); color: var(--pptx-foreground); border: 1px solid var(--pptx-border); border-radius: 8px; box-shadow: 0 8px 24px #0002; font: inherit; font-size: 12px; }
  .host-panel::backdrop { background: #0003; }
  header { display: flex; align-items: center; justify-content: space-between; margin-bottom: 12px; }
  header button { font: inherit; font-size: 20px; border: 0; background: transparent; color: inherit; cursor: pointer; }
</style>
