import { afterEach, describe, expect, it, vi } from 'vitest';

import type { CollaborationConfig } from '../types';
import {
	collaborationWebsocketOptions,
	startCollaborationWebsocket,
} from './collaboration-websocket';

class Provider {
	protocols: string[] = [];
	shouldConnect = false;
	wsconnected = false;
	synced = false;
	listeners = new Map<string, ((...args: never[]) => void)[]>();
	on(event: string, callback: (...args: never[]) => void) {
		this.listeners.set(event, [...(this.listeners.get(event) ?? []), callback]);
	}
	emit(event: string, value?: unknown) {
		this.listeners.get(event)?.forEach((callback) => callback(value as never));
	}
	connect = vi.fn(() => {
		this.shouldConnect = true;
	});
	disconnect = vi.fn(() => {
		this.shouldConnect = false;
		this.emit('connection-close');
	});
	destroy = vi.fn(() => this.disconnect());
}
const config: CollaborationConfig = {
	serverUrl: 'wss://example.test',
	roomId: 'room',
	userName: 'Ada',
};
afterEach(() => vi.useRealTimers());

describe('authenticated collaboration lifecycle', () => {
	it('never configures URL parameters', () => {
		expect(
			collaborationWebsocketOptions({
				...config,
				authToken: 'secret',
				websocketProtocols: ['app.v1'],
			}),
		).not.toHaveProperty('params');
		expect(
			collaborationWebsocketOptions({
				...config,
				authToken: 'secret',
				getWebsocketProtocols: async () => [],
			}),
		).not.toHaveProperty('params');
		expect(collaborationWebsocketOptions(config)).toMatchObject({
			connect: false,
			disableBc: true,
		});
	});

	it('rejects legacy token-only authentication with an explicit migration hint', () => {
		for (const websocketProtocols of [undefined, []]) {
			expect(() =>
				collaborationWebsocketOptions({
					...config,
					authToken: 'private-token',
					websocketProtocols,
				}),
			).toThrow('Use websocketProtocols or getWebsocketProtocols');
		}
	});

	it('waits for new credentials on each reconnect and only reports readiness on sync', async () => {
		vi.useFakeTimers();
		const provider = new Provider();
		const connect = provider.connect;
		let resolve!: (protocols: string[]) => void;
		const getWebsocketProtocols = vi
			.fn()
			.mockResolvedValueOnce(['token.first'])
			.mockImplementationOnce(
				() =>
					new Promise<string[]>((r) => {
						resolve = r;
					}),
			);
		const onstatus = vi.fn();
		startCollaborationWebsocket(provider, {
			...config,
			getWebsocketProtocols,
			onstatus,
		});
		await Promise.resolve();
		expect(provider.protocols).toStrictEqual(['token.first']);
		provider.wsconnected = true;
		provider.emit('status', { status: 'connected' });
		expect(onstatus).not.toHaveBeenCalledWith('synced', undefined);
		provider.emit('sync', true);
		expect(onstatus).toHaveBeenLastCalledWith('synced', undefined);
		provider.wsconnected = false;
		provider.emit('connection-close');
		expect(provider.shouldConnect).toBeFalsy();
		await vi.advanceTimersByTimeAsync(1000);
		expect(getWebsocketProtocols).toHaveBeenCalledTimes(2);
		expect(connect).toHaveBeenCalledOnce();
		resolve(['token.second']);
		await Promise.resolve();
		expect(provider.protocols).toStrictEqual(['token.second']);
		expect(connect).toHaveBeenCalledTimes(2);
		provider.destroy();
	});

	it('reports auth failures, retries, and ignores credentials resolved after teardown', async () => {
		vi.useFakeTimers();
		const provider = new Provider();
		const connect = provider.connect;
		const onstatus = vi.fn();
		let resolve!: (protocols: string[]) => void;
		const getWebsocketProtocols = vi
			.fn()
			.mockRejectedValueOnce(new Error('private credential detail'))
			.mockImplementationOnce(
				() =>
					new Promise<string[]>((r) => {
						resolve = r;
					}),
			);
		startCollaborationWebsocket(provider, {
			...config,
			getWebsocketProtocols,
			onstatus,
		});
		await Promise.resolve();
		expect(onstatus).toHaveBeenLastCalledWith(
			'error',
			new Error('Collaboration authentication or connection failed'),
		);
		await vi.advanceTimersByTimeAsync(1000);
		provider.destroy();
		resolve(['expired']);
		await Promise.resolve();
		await vi.advanceTimersByTimeAsync(60000);
		expect(connect).not.toHaveBeenCalled();
		expect(getWebsocketProtocols).toHaveBeenCalledTimes(2);
	});

	it('keeps an open but unsynced connection unready and reports a timeout', async () => {
		vi.useFakeTimers();
		const provider = new Provider();
		const onstatus = vi.fn();
		startCollaborationWebsocket(provider, { ...config, onstatus });
		provider.wsconnected = true;
		provider.emit('status', { status: 'connected' });
		await vi.advanceTimersByTimeAsync(30000);
		expect(onstatus).toHaveBeenCalledWith(
			'error',
			new Error('Collaboration synchronization timed out'),
		);
		expect(onstatus).not.toHaveBeenCalledWith('synced', undefined);
		provider.destroy();
	});
});
