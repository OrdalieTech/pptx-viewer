/** Svelte collaboration controller. Undo remains local, as in the other bindings. */
import type { PptxSlide } from 'pptx-viewer-core';
import type {
	CollaborationConfig,
	CollabLoadOrigin,
	CollaborationLivePatcher,
	ConnectionStatus,
	RemoteCursor,
	SanitizedPresence,
	YDocLike,
	YjsFactories,
} from 'pptx-viewer-shared';
import {
	createCollaborationLivePatcher,
	createSyncGate,
	registerCollaborationSource,
	createWriteBackScheduler,
	DEFAULT_CURSOR_COLOR,
	isMixedContentBlocked,
	resolveTransportForServerUrl,
	validateRoomId,
} from 'pptx-viewer-shared';

import type { CollaborationDeps } from './collaboration-deps';
import { registerCollaborationEffects } from './collaboration-effects.svelte';
import { CollaborationPresence } from './collaboration-presence.svelte';
import type { CollabProviderHandle } from './collaboration-provider';
import type { ObserveRemoteDeps } from './collaboration-remote-sync';
import {
	adoptDocSlidesAfterLoad,
	observeRemoteSlides,
	publishLocalSlides,
} from './collaboration-remote-sync';
import type { CollabSession, CollabSessionFactory } from './collaboration-session';
import { createDefaultSession } from './collaboration-session';
import { wireInitialSync, wireProviderStatus } from './collaboration-status';

export class CollaborationController {
	/** Live connection status (reactive). */
	status = $state<ConnectionStatus>('disconnected');
	/** Interim Y.Doc channel for in-flight inline text (dormant when stopped). */
	readonly livePatcher: CollaborationLivePatcher = createCollaborationLivePatcher();

	#active = $state(false);
	readonly #deps: CollaborationDeps;
	readonly #makeSession: CollabSessionFactory;

	#session: CollabSession | null = null;
	#ydoc: YDocLike | null = null;
	#factories: YjsFactories | null = null;
	#provider: CollabProviderHandle | null = null;
	#config: CollaborationConfig | null = $state(null);
	#lastStarted: CollaborationConfig | null = null;
	#startedByEffect = false;

	#applyingRemote = false;
	#lastSynced = '';
	#didSync = false;
	#unobserve: (() => void) | null = null;
	#connectTimer: ReturnType<typeof setTimeout> | null = null;

	readonly #gate = createSyncGate(() => {
		if (this.#ydoc && !this.#didSync) {
			adoptDocSlidesAfterLoad(this.#ydoc, this.#remoteDeps());
		}
		this.#didSync = true;
		if (this.#ydoc) {
			this.#flushLocalSlides();
		}
	});
	readonly #writeBack = createWriteBackScheduler({
		getYDoc: () => this.#ydoc,
		getSourceBytes: () => this.#deps.getSourceBytes?.() ?? null,
		getSaveOptions: () => this.#deps.getSaveOptions?.(),
	});
	readonly #presence: CollaborationPresence;

	constructor(deps: CollaborationDeps) {
		this.#deps = deps;
		this.#makeSession = deps.createSession ?? createDefaultSession;
		this.#presence = new CollaborationPresence(() => ({
			width: this.#deps.getCanvasWidth?.(),
			height: this.#deps.getCanvasHeight?.(),
		}));

		registerCollaborationEffects({
			getConfig: () => this.#deps.getConfig(),
			getSlides: () => this.#deps.getSlides(),
			syncConfig: (config) => this.#syncConfig(config),
			isPublishable: () => this.#active && this.#gate.isOpen(),
			flushLocalSlides: (slides) => this.#flushLocalSlides(slides),
			stop: () => this.stop(),
			rejoin: () => {
				if (this.#lastStarted) {
					void this.#run(this.#lastStarted);
				}
			},
		});
	}

	get active(): boolean {
		return this.#active;
	}
	/** Return this session's live Y.Doc, or null when collaboration is stopped. */
	getDocument(): YDocLike | null {
		return this.#ydoc;
	}
	/** Read-only participant (session live with the `viewer` role) - cannot select/drag/mutate. */
	get readOnly(): boolean {
		return this.#active && this.#config?.role === 'viewer';
	}
	get cursors(): RemoteCursor[] {
		return this.#presence.cursors;
	}
	get remotePresences(): SanitizedPresence[] {
		return this.#presence.remotePresences;
	}
	/** Followed peer's client id, or null when free (reactive). */
	get followedClientId(): number | null {
		return this.#presence.followedClientId;
	}
	/** The config the active session was started with (null when stopped); the
	 * Share dialog's active view reads the local user's name/colour from this. */
	get activeCollaboration(): CollaborationConfig | null {
		return this.#config;
	}
	/** Total connected participants (self + remote peers), reactive. */
	get connectedCount(): number {
		return this.remotePresences.length + (this.#active ? 1 : 0);
	}

	/** Publish a cursor move (slide-space px); no-op when no session is active. */
	setCursor(x: number, y: number, activeSlideIndex?: number): void {
		this.#presence.setCursor(x, y, activeSlideIndex);
	}
	/** Publish the local selection; no-op when no session is active. */
	setSelection(selectedElementId: string | undefined, activeSlideIndex?: number): void {
		this.#presence.setSelection(selectedElementId, activeSlideIndex);
	}
	/** Publish the local active-slide index (drives peer follow-along). */
	setActiveSlide(index: number): void {
		this.#presence.setActiveSlide(index);
	}
	/** Follow the given peer's active slide, or `null` to stop following. */
	followUser(clientId: number | null): void {
		this.#presence.followUser(clientId);
	}

	/** Register the parsed source before adopting authoritative room slides. */
	adoptDocAfterLoad(origin: CollabLoadOrigin = 'user'): void {
		if (this.#active && this.#ydoc) {
			registerCollaborationSource(this.#ydoc, this.#deps.getSlides());
			adoptDocSlidesAfterLoad(this.#ydoc, this.#remoteDeps(), origin);
		}
	}

	#syncConfig(config: CollaborationConfig | undefined): void {
		if (config && config !== this.#lastStarted) {
			this.#lastStarted = config;
			this.#startedByEffect = true;
			void this.#run(config);
		} else if (!config && this.#active && this.#startedByEffect) {
			this.#lastStarted = null;
			this.#startedByEffect = false;
			this.stop();
		}
	}

	/** Write the current local slides into the doc (granular, echo-deduped). */
	#flushLocalSlides(slides: PptxSlide[] = this.#deps.getSlides()): void {
		const published = publishLocalSlides({
			slides,
			ydoc: this.#ydoc,
			factories: this.#factories,
			applyingRemote: this.#applyingRemote,
			role: this.#config?.role,
			lastSynced: this.#lastSynced,
		});
		if (published === null) {
			return;
		}
		this.#lastSynced = published;
		if (this.#config) {
			this.#writeBack.schedule(this.#config);
		}
	}
	#clearTimers(): void {
		if (this.#connectTimer !== null) {
			clearTimeout(this.#connectTimer);
			this.#connectTimer = null;
		}
		this.#writeBack.cancel();
	}
	/** Start (or restart) a session with the given config (dialog-driven). */
	async start(config: CollaborationConfig): Promise<void> {
		this.#lastStarted = config;
		this.#startedByEffect = false;
		await this.#run(config);
	}

	#connectionGeneration = 0;

	async #run(config: CollaborationConfig): Promise<void> {
		this.stop();
		const generation = this.#connectionGeneration;
		this.#config = config;
		try {
			validateRoomId(config.roomId);
		} catch {
			this.status = 'error';
			config.onstatus?.('error', new Error('Collaboration unavailable'));
			return;
		}
		const transport = config.transport ?? resolveTransportForServerUrl(config.serverUrl);
		if (transport === 'websocket' && isMixedContentBlocked(config.serverUrl)) {
			this.status = 'error';
			config.onstatus?.('error', new Error('Collaboration unavailable'));
			return;
		}
		this.status = 'connecting';
		try {
			const session = await this.#makeSession(transport, config);
			if (generation !== this.#connectionGeneration) {
				session.destroy();
				return;
			}
			this.#session = session;
			this.#ydoc = session.ydoc;
			registerCollaborationSource(session.ydoc, this.#deps.getSlides());
			this.#factories = session.factories;
			this.livePatcher.configure(session.ydoc, session.factories);
			this.#provider = session.provider;

			wireInitialSync(this.#provider, transport, this.#gate);

			this.#presence.start(this.#provider.awareness, {
				userName: config.userName,
				userColor: config.userColor ?? DEFAULT_CURSOR_COLOR,
				userAvatar: config.userAvatar,
				role: config.role,
			});

			this.#wireProvider(transport, config);

			this.#active = true;
			this.#deps.onStart?.(config);
		} catch (error) {
			if (generation !== this.#connectionGeneration) {
				return;
			}
			this.stop();
			this.status = 'error';
			config.onstatus?.(
				'error',
				error instanceof Error ? error : new Error('Collaboration unavailable'),
			);
		}
	}

	#remoteDeps = (): ObserveRemoteDeps => ({
		onError: (error) => {
			const config = this.#config;
			this.stop();
			this.status = 'error';
			config?.onstatus?.('error', error instanceof Error ? error : new Error(String(error)));
		},
		isApplyingRemote: () => this.#applyingRemote,
		setApplyingRemote: (value) => (this.#applyingRemote = value),
		setLastSynced: (value) => (this.#lastSynced = value),
		applyRemoteSlides: (slides) => this.#deps.applyRemoteSlides(slides),
		scheduleWriteBack: (cfg) => this.#writeBack.schedule(cfg),
	});

	/** Attach the status machine and the remote-slide observer to the session. */
	#wireProvider(transport: string, config: CollaborationConfig): void {
		if (!this.#provider || !this.#ydoc) {
			return;
		}
		this.#unobserve = observeRemoteSlides(this.#ydoc, config, this.#remoteDeps());
		wireProviderStatus(this.#provider, transport, {
			setStatus: (status) => {
				this.status = status;
				if (status === 'error') {
					config.onstatus?.('error', new Error('Collaboration unavailable'));
				}
			},
			getStatus: () => this.status,
			isActive: () => this.#active,
			stop: () => this.stop(),
			gate: this.#gate,
			setConnectTimer: (timer) => (this.#connectTimer = timer),
			getConnectTimer: () => this.#connectTimer,
		});
	}

	stop(): void {
		this.#connectionGeneration++;
		this.#clearTimers();
		this.#gate.reset();
		this.#presence.stop();
		this.#unobserve?.();
		this.#unobserve = null;
		this.#session?.destroy();
		this.#session = null;
		this.#provider = null;
		this.#ydoc = null;
		this.#factories = null;
		this.#config = null;
		this.livePatcher.configure(null, null);
		this.#applyingRemote = false;
		this.#lastSynced = '';
		this.#didSync = false;
		if (this.#active) {
			this.#deps.onStop?.();
		}
		this.#active = false;
		this.status = 'disconnected';
	}
}
