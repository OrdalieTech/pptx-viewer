import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';

import { expect, test } from 'vitest';
import { Doc } from 'yjs';

import { PptxCodec } from '../../codec/index.js';
import {
	fixtureBytes,
	fixtureMissing,
	replaceText,
	restoredDoc,
	slideMaps,
} from './fidelity-helpers.js';

function available(command: string): boolean {
	const result = spawnSync(command, ['--version'], { timeout: 10_000 });
	// A broken installed tool is a failure, not an absent dependency.
	if (result.error && 'code' in result.error && result.error.code === 'ENOENT') {
		return false;
	}
	if (result.error) {
		throw result.error;
	}
	return true;
}

function run(command: string, args: string[]): string {
	const result = spawnSync(command, args, { encoding: 'utf8', timeout: 60_000 });
	if (result.error) {
		throw result.error;
	}
	expect(result.status, `${command}: ${result.stderr}\n${result.stdout}`).toBe(0);
	return result.stdout;
}

const soffice = process.env.LIBREOFFICE_BIN || 'soffice';
const canRender = available(soffice) && available('pdftoppm') && available('pdfinfo');

test.skipIf(!canRender || fixtureMissing('sample-deck.pptx'))(
	'libreOffice renders a persisted text edit while unedited slides remain pixel-identical',
	async () => {
		const directory = await mkdtemp(join(tmpdir(), 'pptx-codec-render-'));
		const doc = new Doc();
		let restored: Doc | undefined;
		try {
			const source = await fixtureBytes('sample-deck.pptx');
			const codec = new PptxCodec();
			await codec.hydrate(doc, source);
			for (const slide of slideMaps(doc)) {
				slide.set('isDirty', false);
			}
			replaceText(slideMaps(doc)[4], 'Persisted collaborative edit');
			restored = restoredDoc(doc);
			const output = await codec.dehydrate(restored, source);
			await writeFile(join(directory, 'source.pptx'), source);
			await writeFile(join(directory, 'edited.pptx'), output);
			const profile = pathToFileURL(join(directory, 'profile')).href;
			for (const name of ['source', 'edited']) {
				run(soffice, [
					`-env:UserInstallation=${profile}`,
					'--headless',
					'--convert-to',
					'pdf',
					'--outdir',
					directory,
					join(directory, `${name}.pptx`),
				]);
				const info = run('pdfinfo', [join(directory, `${name}.pdf`)]);
				expect(info).toMatch(/Pages:\s+7\b/);
				for (const page of [1, 5, 7]) {
					run('pdftoppm', [
						'-f',
						String(page),
						'-l',
						String(page),
						'-singlefile',
						'-scale-to',
						'800',
						join(directory, `${name}.pdf`),
						join(directory, `${name}-${page}`),
					]);
				}
			}
			for (const page of [1, 5, 7]) {
				const before = await readFile(join(directory, `source-${page}.ppm`));
				const after = await readFile(join(directory, `edited-${page}.ppm`));
				const digest = (bytes: Uint8Array) => createHash('sha256').update(bytes).digest('hex');
				if (page === 5) {
					expect(digest(after), 'edited slide must visibly change').not.toBe(digest(before));
				} else {
					expect(digest(after), `unedited slide ${page} must render identically`).toBe(
						digest(before),
					);
				}
			}
		} finally {
			doc.destroy();
			restored?.destroy();
			await rm(directory, { recursive: true, force: true });
		}
	},
	180_000,
);
