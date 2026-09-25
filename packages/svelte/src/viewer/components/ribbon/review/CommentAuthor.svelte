<script lang="ts">
	import { getContext, type Snippet } from 'svelte';
	const { author, createdAt }: { author: string; createdAt?: string } = $props();
	const hostAvatar = getContext<(() => Snippet<[string]> | undefined) | undefined>('pptx-comment-avatar');
	const avatar = $derived(hostAvatar?.());
	const initials = $derived(author.trim().split(/\s+/).slice(0, 2).map((part) => part[0]).join('').toLocaleUpperCase() || '?');
	const timestamp = $derived(createdAt ? Date.parse(createdAt) : NaN);
	const relativeTime = $derived.by(() => {
		if (!Number.isFinite(timestamp)) return '';
		const seconds = Math.round((timestamp - Date.now()) / 1000);
		const units: [number, Intl.RelativeTimeFormatUnit][] = [
			[31_536_000, 'year'], [2_592_000, 'month'], [86_400, 'day'],
			[3_600, 'hour'], [60, 'minute'], [1, 'second'],
		];
		const [duration, unit] = units.find(([duration]) => Math.abs(seconds) >= duration) ?? units.at(-1)!;
		return new Intl.RelativeTimeFormat(undefined, { numeric: 'auto' }).format(Math.round(seconds / duration), unit);
	});
</script>

<div class="pptx-svelte-comment-author">
	{#if avatar}
		{@render avatar(author)}
	{:else}
		<span class="pptx-svelte-comment-avatar" aria-hidden="true">{initials}</span>
	{/if}
	<div class="pptx-svelte-comment-author-details">
		<strong title={author}>{author}</strong>
		{#if relativeTime}<time datetime={createdAt} title={new Date(timestamp).toLocaleString()}>{relativeTime}</time>{/if}
	</div>
</div>

<style>
	.pptx-svelte-comment-author { display: flex; min-width: 0; align-items: center; gap: 8px; }
	.pptx-svelte-comment-avatar { display: grid; flex: 0 0 32px; place-items: center; width: 32px; height: 32px; border-radius: 50%; background: var(--pptx-muted, #e2e8f0); color: #344252; font-size: 12px; font-weight: 500; }
	.pptx-svelte-comment-author-details { display: grid; min-width: 0; gap: 3px; }
	.pptx-svelte-comment-author-details strong { overflow: hidden; color: var(--pptx-card-foreground, inherit); font-size: 13px; font-weight: 600; line-height: 1.1; text-overflow: ellipsis; white-space: nowrap; }
	.pptx-svelte-comment-author-details time { color: var(--pptx-muted-foreground, #94a3b8); font-size: 11px; line-height: 1.1; }
</style>
