import type { CollaborationConfig } from '../types';
import { CONNECTION_TIMEOUT_MS } from './collaboration-presence';

const sessionFailures = new WeakMap<object, (error: Error) => void>();
/** Stop the transport before reporting an invalid authoritative document. */
export function failCollaborationSession(doc: object, error: Error): void {
	sessionFailures.get(doc)?.(error);
}

/** The y-websocket surface needed to refresh one-use credentials safely. */
export interface AuthenticatedWebsocketProvider {
	protocols: string[];
	shouldConnect: boolean;
	wsconnected: boolean;
	synced: boolean;
	connect(): void;
	disconnect(): void;
	destroy(): void;
	on(event: 'status', callback: (event: { status: string }) => void): void;
	on(event: 'sync', callback: (synced: boolean) => void): void;
	on(event: 'connection-close' | 'connection-error', callback: (event?: { code?: number }) => void): void;
}

/** Start only after listeners and fresh credentials are installed. Server rooms never sync via BroadcastChannel. */
export function collaborationWebsocketOptions(config: CollaborationConfig) {
	if (config.authToken && !config.websocketProtocols?.length && !config.getWebsocketProtocols) {
		throw new Error(
			'WebSocket authToken is no longer supported. Use websocketProtocols or getWebsocketProtocols.',
		);
	}
	return {
		connect: false,
		disableBc: true,
		protocols: config.websocketProtocols,
	};
}

/** Own reconnect scheduling so y-websocket cannot race an asynchronous token refresh. */
export function startCollaborationWebsocket(
	provider: AuthenticatedWebsocketProvider,
	config: CollaborationConfig,
	doc?: object,
): void {
	let stopped = false;
	let pending = false;
	let generation = 0;
	let retryTimer: ReturnType<typeof setTimeout> | undefined;
	let syncTimer: ReturnType<typeof setTimeout> | undefined;
	let attempts = 0;
	const connect = provider.connect.bind(provider);
	const disconnect = provider.disconnect.bind(provider);
	const destroy = provider.destroy.bind(provider);
	const report = (
		status: Parameters<NonNullable<CollaborationConfig['onstatus']>>[0],
		error?: Error,
	) => {
		if (!stopped) {
			config.onstatus?.(status, error);
		}
	};
	const schedule = () => {
		provider.shouldConnect = false;
		if (stopped || retryTimer) {
			return;
		}
		retryTimer = setTimeout(
			() => {
				retryTimer = undefined;
				void attempt();
			},
			Math.min(1000 * 2 ** attempts++, 10000),
		);
	};
	const attempt = async () => {
		if (stopped || pending) {
			return;
		}
		pending = true;
		const current = generation;
		report('connecting');
		clearTimeout(syncTimer);
		syncTimer = setTimeout(() => {
			if (stopped || provider.synced) {
				return;
			}
			report('error', new Error('Collaboration synchronization timed out'));
			generation++;
			pending = false;
			disconnect();
			schedule();
		}, CONNECTION_TIMEOUT_MS);
		try {
			const protocols = config.getWebsocketProtocols
				? await config.getWebsocketProtocols()
				: config.websocketProtocols;
			if (stopped || current !== generation) {
				return;
			}
			if (protocols) {
				provider.protocols = protocols;
			}
			connect();
			// All retries pass through attempt(), including failed handshakes.
			provider.shouldConnect = false;
		} catch {
			if (stopped || current !== generation) {
				return;
			}
			clearTimeout(syncTimer);
			report('error', new Error('Collaboration authentication or connection failed'));
			schedule();
		} finally {
			if (current === generation) {
				pending = false;
			}
		}
	};
	provider.on('status', ({ status }) => {
		if (status === 'connected') {
			report('connected');
		}
		if (status === 'disconnected') {
			report('disconnected');
		}
	});
	provider.on('sync', (synced) => {
		if (synced && provider.wsconnected) {
			clearTimeout(syncTimer);
			attempts = 0;
			report('synced');
		}
	});
	provider.on('connection-error', () =>
		report('error', new Error('Collaboration connection failed')),
	);
	provider.on('connection-close', (event) => {
		if (event?.code === 4409 || event?.code === 4413) {
			clearTimeout(syncTimer);
			provider.disconnect();
			config.onstatus?.('error', new Error(event.code === 4409
				? 'Un conflit avec le fichier externe empêche la sauvegarde. Les modifications ne sont pas sauvegardées.'
				: 'La présentation dépasse la limite de sauvegarde externe. Les modifications ne sont pas sauvegardées.'));
			return;
		}
		clearTimeout(syncTimer);
		report('disconnected');
		schedule();
	});
	const stop = () => {
		stopped = true;
		generation++;
		clearTimeout(retryTimer);
		clearTimeout(syncTimer);
	};
	provider.connect = () => {
		stopped = false;
		void attempt();
	};
	provider.disconnect = () => {
		stop();
		disconnect();
	};
	provider.destroy = () => {
		stop();
		destroy();
	};
	if (doc) {
		sessionFailures.set(doc, (error) => {
			provider.disconnect();
			config.onstatus?.('error', error);
		});
	}
	void attempt();
}
