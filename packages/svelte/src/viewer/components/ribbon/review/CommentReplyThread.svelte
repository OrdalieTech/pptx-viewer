<script lang="ts">
	/**
	 * CommentReplyThread: a comment's `replies` list, rendered RECURSIVELY
	 * (wave-4 B5). `PptxComment.replies` is itself `PptxComment[]`, so a reply
	 * can carry its own nested replies; core now nests legacy `p:cmLst` replies
	 * the same way it always nested modern ones, but the previous panel only
	 * ever unrolled ONE level (`{#each comment.replies as reply}` with no
	 * recursion into `reply.replies`), so a legacy comment loaded with a
	 * grandchild reply silently dropped it from view. Self-recursive component,
	 * the same pattern `ElementRenderer` uses for group children.
	 */
	import type { PptxComment } from 'pptx-viewer-core';

	import { useTranslator } from '../../../../i18n/context';
	import CommentBody from '../../CommentBody.svelte';
	import CommentAuthor from './CommentAuthor.svelte';
	// Self-import: a reply's own replies recurse into this same component.
	// eslint-disable-next-line import/no-self-import
	import CommentReplyThread from './CommentReplyThread.svelte';

	const { replies }: { replies: readonly PptxComment[] } = $props();
	const t = useTranslator();
</script>

{#if replies.length > 0}
	<div class="pptx-svelte-comment-replies">
		{#each replies as reply (reply.id)}
			<div class="pptx-svelte-comment-reply">
				<CommentAuthor author={reply.author ?? t('pptx.comments.unknownAuthor')} createdAt={reply.createdAt} />
				<p><CommentBody text={reply.text} mentions={reply.mentions} /></p>
				{#if reply.replies && reply.replies.length > 0}
					<CommentReplyThread replies={reply.replies} />
				{/if}
			</div>
		{/each}
	</div>
{/if}

<style>
	.pptx-svelte-comment-replies {
		display: grid;
		gap: 10px;
		margin-left: 20px;
		padding-left: 12px;
		border-left: 1px solid var(--pptx-border, #33334d);
	}
	.pptx-svelte-comment-reply { display: grid; gap: 5px; }
	.pptx-svelte-comment-reply p {
		margin: 0;
		padding-left: 32px;
		font-size: 12px;
		line-height: 1.45;
		white-space: pre-wrap;
		overflow-wrap: anywhere;
	}
</style>
