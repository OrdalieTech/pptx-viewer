<script lang="ts">
 import { elementTypeLabelKey } from "pptx-viewer-shared";
	/**
	 * ReviewCommentsPanel: history-aware comment review for the active slide.
	 * Comment list transforms are shared with the other framework bindings; this
	 * component only owns the compact review UI and writes through EditorState.
	 *
	 * Wave-4 B5: both composers (new comment, reply) offer an `@`-mention
	 * typeahead ({@link CommentComposeState} + {@link CommentMentionSuggestions}),
	 * and replies render recursively ({@link CommentReplyThread}) so a legacy
	 * `p:cmLst` reply nested more than one level deep is no longer dropped.
	 *
	 * `addCommentToList` / `replyToCommentInList` (shared) build the new
	 * comment/reply but do not accept a `mentions` array, so `addComment` /
	 * `submitReply` fold the accepted mentions onto the newly-appended row
	 * afterwards: both helpers append at a known position (the list's last
	 * item, and the parent's last reply respectively).
	 */
	import { getContext } from 'svelte';
	import type { PptxComment, PptxCommentMention, PptxModernCommentAuthor, PptxSlide } from 'pptx-viewer-core';
	import {
		addCommentToList,
		removeCommentFromList,
		replyToCommentInList,
		toggleCommentResolvedInList,
	} from 'pptx-viewer-shared';

	import { useTranslator } from '../../../../i18n/context';
	import { useViewerOptions } from '../../../state/viewer-options-context';
	import type { EditorState } from '../../../editor/editor-state.svelte';
	import { CommentComposeState } from './comment-compose.svelte';
	import CommentCard from './CommentCard.svelte';
	import CommentMentionSuggestions from './CommentMentionSuggestions.svelte';

	const {
		editor,
		embedded = false,
	}: {
		editor: EditorState;
		/**
		 * True when hosted inside a chrome that already renders its own title +
		 * close button (the mobile `MobileSheet`). Suppresses the internal
		 * heading so mobile doesn't show "Comments" twice stacked.
		 */
		embedded?: boolean;
	} = $props();
	const t = useTranslator();
	const optionsState = useViewerOptions();
	const collaboratorName = getContext<(() => string | undefined) | undefined>('pptx-comment-author');
	const authorName = $derived(
		collaboratorName?.()?.trim() ||
		optionsState.options.general.userName ||
		t('pptx.comments.defaultAuthorName'),
	);
	const compose = new CommentComposeState();
	const slide = $derived(editor.slides[editor.currentSlideIndex]);
	const comments = $derived(slide?.comments ?? []);
	const selectedLabel = $derived(editor.selectedElement ? t(elementTypeLabelKey(editor.selectedElement.type)) : null);
	/**
	 * The `@`-mention author catalogue: modern authors plus legacy
	 * (`ppt/commentAuthors.xml`) ones mapped into the same shape (their `id`
	 * doubles as `userId`; `providerId` is a fixed marker, unused by matching
	 * or by `insertCommentMention`, which only reads `id`/`name`).
	 */
	const mentionAuthors = $derived<PptxModernCommentAuthor[]>([
		...editor.modernCommentAuthors,
		...editor.commentAuthors.map((author) => ({
			id: author.id,
			name: author.name,
			initials: author.initials,
			userId: author.id,
			providerId: 'legacy',
		})),
	]);

	function replaceComments(next: PptxComment[]): void {
		const index = editor.currentSlideIndex;
		const slides = editor.slides.map((item, itemIndex) =>
			itemIndex === index ? { ...item, comments: next } : item,
		) as PptxSlide[];
		editor.commitSlides(slides);
	}

	function addComment(): void {
		const next = addCommentToList(comments, compose.text, authorName);
		if (!next) {
			return;
		}
		replaceComments(
			compose.mentions.length > 0
				? next.map((c, i) => (i === next.length - 1 ? { ...c, mentions: compose.mentions } : c))
				: next,
		);
		compose.reset();
	}

	function toggleResolved(id: string): void {
		const next = toggleCommentResolvedInList(comments, id);
		if (next) {
			replaceComments(next);
		}
	}

	function removeComment(id: string): void {
		const next = removeCommentFromList(comments, id);
		if (next) {
			replaceComments(next);
		}
	}

	function submitReply(id: string, text: string, mentions: PptxCommentMention[]): boolean {
		const next = replyToCommentInList(comments, id, text, authorName);
		if (!next) {
			return false;
		}
		replaceComments(
			mentions.length > 0
				? next.map((c) => {
						const parentReplies = c.replies;
						if (c.id !== id || !parentReplies || parentReplies.length === 0) {
							return c;
						}
						const replies = parentReplies.map((r, i) =>
							i === parentReplies.length - 1 ? { ...r, mentions } : r,
						);
						return { ...c, replies };
					})
				: next,
		);
		return true;
	}

	// Wire a composer's `<textarea>` to its `CommentComposeState` for the
	// typeahead: caret tracking on input/click/arrow-keys, and the suggestion
	// list's keyboard navigation. Two small factories (not one that returns a
	// handler bag) so each stays a plain, directly-typed element attribute --
	// the idiom this codebase uses everywhere else for a typed `currentTarget`.
	function onComposeChange(state: CommentComposeState) {
		return (event: Event & { currentTarget: HTMLTextAreaElement }) => {
			state.onInput(event.currentTarget.value, event.currentTarget.selectionStart ?? 0);
		};
	}
	function onComposeKeydown(state: CommentComposeState) {
		return (event: KeyboardEvent & { currentTarget: HTMLTextAreaElement }) => {
			const result = state.onKeydown(event, mentionAuthors);
			if (!result.consumed) {
				return;
			}
			event.preventDefault();
			if (result.caret === undefined) {
				return;
			}
			const target = event.currentTarget;
			const caret = result.caret;
			queueMicrotask(() => target.setSelectionRange(caret, caret));
		};
	}
</script>

<section
	class="pptx-svelte-comments"
	aria-label={embedded ? t('pptx.comments.slideComments') : undefined}
	aria-labelledby={embedded ? undefined : 'pptx-svelte-comments-title'}
>
	{#if !embedded}
		<div class="pptx-svelte-comments-heading">
			<div>
				<span class="pptx-svelte-comments-eyebrow">{t('pptx.ribbon.tab.review')}</span>
				<h3 id="pptx-svelte-comments-title">{t('pptx.comments.slideComments')}</h3>
			</div>
			<span class="pptx-svelte-comments-count">{comments.length}</span>
		</div>
	{/if}

	{#if selectedLabel}
		<p class="pptx-svelte-comments-target">{t('pptx.comments.commentingOn')} {selectedLabel}</p>
	{/if}
	<div class="pptx-svelte-comments-compose">
		<div class="pptx-svelte-comment-composer">
			<textarea
				value={compose.text}
				rows="2"
				placeholder={t('pptx.comments.mentionPlaceholder')}
				aria-label={t('pptx.comments.addComment')}
				oninput={onComposeChange(compose)}
				onkeyup={onComposeChange(compose)}
				onclick={onComposeChange(compose)}
				onkeydown={onComposeKeydown(compose)}
			></textarea>
			<CommentMentionSuggestions
				authors={compose.suggestions(mentionAuthors)}
				highlightIndex={compose.highlightIndex}
				onselect={(author) => compose.accept(author)}
			/>
		</div>
		<button type="button" disabled={!compose.text.trim()} onclick={addComment}>{t('pptx.comments.addComment')}</button>
	</div>

	{#if comments.length === 0}
		<p class="pptx-svelte-comments-empty">{t('pptx.comments.noneOnSlide')}</p>
	{:else}
		<div class="pptx-svelte-comments-list" aria-label={t('pptx.comments.slideComments')}>
			{#each comments as comment (comment.id)}
				<CommentCard {comment} {mentionAuthors} onResolve={toggleResolved} onRemove={removeComment} onReply={submitReply} />
			{/each}
		</div>
	{/if}
</section>

<style>
	.pptx-svelte-comments { display: grid; gap: 8px; width: min(340px, 100%); padding-left: 12px; border-left: 1px solid var(--pptx-border, #33334d); }
	.pptx-svelte-comments-heading { display: flex; justify-content: space-between; align-items: flex-start; gap: 10px; }
	.pptx-svelte-comments-eyebrow { display: block; color: var(--pptx-muted-foreground, #94a3b8); font-size: 10px; font-weight: 700; letter-spacing: .08em; text-transform: uppercase; }
	.pptx-svelte-comments h3 { margin: 1px 0 0; font-size: 13px; }
	.pptx-svelte-comments-count { display: grid; place-items: center; min-width: 20px; height: 20px; border-radius: 10px; background: var(--pptx-muted, #2a2a3d); font-size: 11px; font-weight: 700; }
	.pptx-svelte-comments-target, .pptx-svelte-comments-empty { margin: 0; color: var(--pptx-muted-foreground, #94a3b8); font-size: 11px; }
	.pptx-svelte-comments-compose { display: grid; gap: 5px; }
	.pptx-svelte-comment-composer { position: relative; }
	.pptx-svelte-comments-compose textarea { box-sizing: border-box; width: 100%; resize: vertical; padding: 8px; border: 1px solid var(--pptx-border, #33334d); border-radius: 8px; background: var(--pptx-card, #1e1e2e); color: inherit; font: inherit; font-size: 12px; }
	.pptx-svelte-comments-compose button { justify-self: end; padding: 4px 8px; border: 0; border-radius: var(--pptx-radius, 6px); background: var(--pptx-primary, #6366f1); color: var(--pptx-primary-foreground, white); cursor: pointer; font: inherit; font-size: 11px; font-weight: 600; }
	.pptx-svelte-comments-compose button:disabled { cursor: default; opacity: .45; }
	.pptx-svelte-comments-list { display: grid; gap: 8px; max-height: 300px; overflow-y: auto; }
</style>
