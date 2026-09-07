import type { CollaborationConfig } from 'pptx-viewer-shared';
// @vitest-environment jsdom
import { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import * as Y from 'yjs';

import { createWebsocketBundle as angular } from '../../../../../angular/src/viewer/collaboration-providers';
import { createCollabProvider as svelte } from '../../../../../svelte/src/viewer/collab/collaboration-provider';
import { createCollabProvider as vanilla } from '../../../../../vanilla/src/viewer/collab/collaboration-provider';
import { createCollabProvider as vue } from '../../../../../vue/src/viewer/composables/collaboration-provider';
import { useYjsProvider } from './useYjsProvider';

vi.mock(
	import('../../../../../angular/src/internal/shared'),
	async () => await import('pptx-viewer-shared'),
);

class Socket {
	static sockets: Socket[] = [];
	OPEN = 1;
	readyState = 0;
	binaryType = '';
	onopen?: () => void;
	onclose?: (event: object) => void;
	onerror?: (event: object) => void;
	onmessage?: (event: { data: ArrayBuffer }) => void;
	constructor(
		public url: string,
		public protocols: string[],
	) {
		Socket.sockets.push(this);
	}
	send() {}
	close() {
		this.readyState = 3;
	}
	open() {
		this.readyState = 1;
		this.onopen?.();
	}
	drop() {
		this.readyState = 3;
		this.onclose?.({ code: 1006 });
	}
}

let dispose: (() => void | Promise<void>) | undefined;
beforeEach(() => {
	vi.useFakeTimers();
	Socket.sockets = [];
	vi.stubGlobal('WebSocket', Socket);
	vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true);
});
afterEach(async () => {
	await dispose?.();
	dispose = undefined;
	vi.useRealTimers();
	vi.unstubAllGlobals();
});

async function startReact(config: CollaborationConfig): Promise<() => Promise<void>> {
	const target = document.createElement('div');
	document.body.append(target);
	let root: Root;
	function Host() {
		useYjsProvider({ config });
		return null;
	}
	await act(async () => {
		root = createRoot(target);
		root.render(<Host />);
	});
	await act(async () => {
		await vi.waitFor(() => {
			if (config.authToken && !config.websocketProtocols?.length && !config.getWebsocketProtocols) {
				expect(config.onstatus).toHaveBeenCalledWith(
					'error',
					expect.objectContaining({
						message: expect.stringContaining('Use websocketProtocols or getWebsocketProtocols'),
					}),
				);
			} else {
				expect(Socket.sockets).toHaveLength(1);
			}
		});
	});
	return async () => {
		await act(async () => root.unmount());
		target.remove();
	};
}

const bindings = {
	svelte: async (config: CollaborationConfig) => {
		const doc = new Y.Doc();
		const provider = await svelte('websocket', config, doc);
		return () => {
			provider.destroy();
			doc.destroy();
		};
	},
	vue: async (config: CollaborationConfig) => {
		const doc = new Y.Doc();
		const provider = await vue('websocket', config, doc);
		return () => {
			provider.destroy();
			doc.destroy();
		};
	},
	vanilla: async (config: CollaborationConfig) => {
		const doc = new Y.Doc();
		const provider = await vanilla('websocket', config, doc);
		return () => {
			provider.destroy();
			doc.destroy();
		};
	},
	angular: async (config: CollaborationConfig) => {
		const bundle = await angular(config);
		return () => {
			bundle.departure.dispose();
			bundle.provider.destroy();
			bundle.doc.destroy();
		};
	},
	react: startReact,
};

describe.each(Object.entries(bindings))('%s websocket authentication', (_name, start) => {
	it('rejects legacy token-only authentication before constructing any socket URL', async () => {
		const config: CollaborationConfig = {
			serverUrl: 'wss://collab.test/documents/doc-1',
			roomId: 'ws',
			userName: 'Ada',
			authToken: 'private-token',
			onstatus: vi.fn(),
		};
		if (_name === 'react') {
			dispose = await start(config);
		} else {
			await expect(start(config)).rejects.toThrow(
				'Use websocketProtocols or getWebsocketProtocols',
			);
		}
		expect(Socket.sockets).toHaveLength(0);
	});

	it('uses protocols exclusively, gets a fresh token after a drop, and reports sync separately from socket open', async () => {
		const getWebsocketProtocols = vi
			.fn()
			.mockResolvedValueOnce(['app.v1', 'token.first'])
			.mockResolvedValueOnce(['app.v1', 'token.second']);
		const onstatus = vi.fn();
		dispose = await start({
			serverUrl: 'wss://collab.test/documents/doc-1',
			roomId: 'ws',
			userName: 'Ada',
			authToken: 'must-not-leak',
			websocketProtocols: ['app.v1'],
			getWebsocketProtocols,
			onstatus,
		});
		await act(async () => {
			await vi.advanceTimersByTimeAsync(0);
		});
		expect(Socket.sockets).toHaveLength(1);
		expect(Socket.sockets[0].url).toBe('wss://collab.test/documents/doc-1/ws');
		expect(Socket.sockets[0].protocols).toStrictEqual(['app.v1', 'token.first']);
		await act(async () => Socket.sockets[0].open());
		expect(onstatus).toHaveBeenCalledWith('connected', undefined);
		expect(onstatus).not.toHaveBeenCalledWith('synced', undefined);
		// y-websocket sync message (0), syncStep2 (1), empty Yjs update length (2), update [0,0].
		await act(async () =>
			Socket.sockets[0].onmessage?.({
				data: new Uint8Array([0, 1, 2, 0, 0]).buffer,
			}),
		);
		expect(onstatus).toHaveBeenCalledWith('synced', undefined);
		await act(async () => {
			Socket.sockets[0].drop();
			await vi.advanceTimersByTimeAsync(1000);
		});
		expect(Socket.sockets).toHaveLength(2);
		expect(Socket.sockets[1].protocols).toStrictEqual(['app.v1', 'token.second']);
		expect(Socket.sockets[1].url).toBe('wss://collab.test/documents/doc-1/ws');
		expect(getWebsocketProtocols).toHaveBeenCalledTimes(2);
	});
});
