<script lang="ts">
	import type { PptxComment, PptxCommentMention, PptxModernCommentAuthor } from 'pptx-viewer-core';
	import { useTranslator } from '../../../../i18n/context';
	import CommentBody from '../../CommentBody.svelte';
	import CommentAuthor from './CommentAuthor.svelte';
	import { CommentComposeState } from './comment-compose.svelte';
	import CommentMentionSuggestions from './CommentMentionSuggestions.svelte';
	import CommentReplyThread from './CommentReplyThread.svelte';

	const {
		comment,
		mentionAuthors,
		onResolve,
		onRemove,
		onReply,
	}: {
		comment: PptxComment;
		mentionAuthors: readonly PptxModernCommentAuthor[];
		onResolve: (id: string) => void;
		onRemove: (id: string) => void;
		onReply: (id: string, text: string, mentions: PptxCommentMention[]) => boolean;
	} = $props();
	const t = useTranslator();
	const reply = new CommentComposeState();

	function updateReply(event: Event & { currentTarget: HTMLTextAreaElement }): void {
		reply.onInput(event.currentTarget.value, event.currentTarget.selectionStart ?? 0);
	}

	function submitReply(): void {
		if (reply.text.trim() && onReply(comment.id, reply.text, reply.mentions)) reply.reset();
	}

	function onReplyKeydown(event: KeyboardEvent & { currentTarget: HTMLTextAreaElement }): void {
		const result = reply.onKeydown(event, mentionAuthors);
		if (result.consumed) {
			event.preventDefault();
			if (result.caret !== undefined) {
				const target = event.currentTarget;
				queueMicrotask(() => target.setSelectionRange(result.caret!, result.caret!));
			}
		} else if (event.key === 'Enter' && (event.metaKey || event.ctrlKey)) {
			event.preventDefault();
			submitReply();
		}
	}
</script>

<article class:resolved={comment.resolved} class="pptx-svelte-comment-card">
	<div class="pptx-svelte-comment-meta">
		<CommentAuthor author={comment.author ?? t('pptx.comments.unknownAuthor')} createdAt={comment.createdAt} />
		{#if comment.resolved}<span class="pptx-svelte-comment-resolved">{t('pptx.comments.resolved')}</span>{/if}
	</div>
	<p class="pptx-svelte-comment-body"><CommentBody text={comment.text} mentions={comment.mentions} /></p>
	<CommentReplyThread replies={comment.replies ?? []} />
	<div class="pptx-svelte-comment-reply-compose">
		<div class="pptx-svelte-comment-composer">
			<textarea
				value={reply.text}
				rows="1"
				placeholder={`${t('pptx.comments.reply')}…`}
				aria-label={t('pptx.comments.reply')}
				oninput={updateReply}
				onclick={updateReply}
				onkeyup={(event) => { if (!(event.key === 'Enter' && (event.metaKey || event.ctrlKey))) updateReply(event); }}
				onkeydown={onReplyKeydown}
			></textarea>
			<CommentMentionSuggestions
				authors={reply.suggestions(mentionAuthors)}
				highlightIndex={reply.highlightIndex}
				onselect={(author) => reply.accept(author)}
			/>
		</div>
		{#if reply.text.trim()}
			<button type="button" class="pptx-svelte-comment-reply-submit" onclick={submitReply}>{t('pptx.comments.reply')}</button>
		{/if}
	</div>
	<div class="pptx-svelte-comment-actions">
		<button type="button" class="remove" onclick={() => onRemove(comment.id)}>{t('pptx.comments.remove')}</button>
		<button type="button" class="resolve" onclick={() => onResolve(comment.id)}>{comment.resolved ? t('pptx.comments.reopen') : t('pptx.comments.resolve')}</button>
	</div>
</article>

<style>
	.pptx-svelte-comment-card { display: grid; gap: 10px; padding: 12px; border: 1px solid var(--pptx-border, #33334d); border-radius: 10px; background: var(--pptx-card, #1e1e2e); color: var(--pptx-card-foreground, inherit); box-shadow: 0 1px 4px rgb(0 0 0 / .08); }
	.pptx-svelte-comment-meta { display: flex; align-items: flex-start; justify-content: space-between; gap: 8px; }
	.pptx-svelte-comment-resolved { border-radius: 99px; padding: 2px 7px; background: #dcfce7; color: #166534; font-size: 10px; font-weight: 600; }
	.pptx-svelte-comment-body { margin: 0; font-size: 14px; line-height: 1.45; white-space: pre-wrap; overflow-wrap: anywhere; }
	.pptx-svelte-comment-reply-compose { display: grid; justify-items: end; gap: 6px; padding-top: 10px; border-top: 1px solid var(--pptx-border, #33334d); }
	.pptx-svelte-comment-composer { position: relative; width: 100%; }
	.pptx-svelte-comment-composer textarea { box-sizing: border-box; width: 100%; min-height: 36px; resize: vertical; padding: 8px; border: 1px solid var(--pptx-border, #33334d); border-radius: 8px; background: var(--pptx-card, #1e1e2e); color: inherit; font: inherit; font-size: 12px; }
	.pptx-svelte-comment-composer textarea::placeholder { color: var(--pptx-muted-foreground, #94a3b8); }
	.pptx-svelte-comment-composer textarea:focus-visible { outline: 2px solid var(--pptx-ring, #818cf8); outline-offset: 1px; }
	.pptx-svelte-comment-reply-submit, .pptx-svelte-comment-actions button { border: 0; border-radius: 6px; padding: 4px 7px; background: transparent; cursor: pointer; font: inherit; font-size: 11px; }
	.pptx-svelte-comment-reply-submit, .pptx-svelte-comment-actions .resolve { color: #15803d; }
	.pptx-svelte-comment-actions { display: flex; justify-content: flex-end; gap: 8px; }
	.pptx-svelte-comment-actions .remove { color: var(--pptx-muted-foreground, #94a3b8); }
	.pptx-svelte-comment-actions button:hover, .pptx-svelte-comment-reply-submit:hover { background: var(--pptx-accent, #33334d); }
	.pptx-svelte-comment-actions button:focus-visible, .pptx-svelte-comment-reply-submit:focus-visible { outline: 2px solid var(--pptx-ring, #818cf8); }
</style>
