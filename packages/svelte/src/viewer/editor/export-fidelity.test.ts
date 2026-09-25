import JSZip from 'jszip';
import { PptxHandler } from 'pptx-viewer-core';
import { expect, it, vi } from 'vitest';

import { saveEditorDocument } from './editor-document-state';
import type { EditorSnapshot } from './editor-document-state';

const snapshot: EditorSnapshot = {
	slides: [],
	templateElementsBySlideId: {},
	slideMasters: [],
	notesMaster: undefined,
	handoutMaster: undefined,
	sections: [],
	headerFooter: {},
	presentationProperties: {},
	customShows: [],
	coreProperties: undefined,
	appProperties: undefined,
	customProperties: [],
	tagCollections: [],
};

it('blocks a browser save when serialization reports lost fidelity', async () => {
	const save = vi.fn().mockResolvedValue(new Uint8Array([1, 2, 3]));
	const handler = {
		save,
		getCompatibilityWarnings: () =>
			save.mock.calls.length
				? [{ scope: 'save', code: 'SIGNATURE_LOST', message: 'Signature cannot be preserved' }]
				: [],
	} as unknown as PptxHandler;
	await expect(saveEditorDocument(handler, snapshot)).rejects.toThrow('SIGNATURE_LOST');
	expect(save).toHaveBeenCalledOnce();
});

it('still blocks the second export of a signed presentation', async () => {
	const { handler: source, data } = await PptxHandler.create({ initialSlideCount: 1 });
	const zip = await JSZip.loadAsync(await source.save(data.slides));
	source.dispose();
	zip.file('_xmlsignatures/sig1.xml', '<Signature xmlns="http://www.w3.org/2000/09/xmldsig#"/>');
	const signed = await zip.generateAsync({ type: 'uint8array' });
	const handler = new PptxHandler();
	try {
		const loaded = await handler.load(signed.buffer as ArrayBuffer);
		const signedSnapshot = { ...snapshot, slides: loaded.slides };
		await expect(saveEditorDocument(handler, signedSnapshot)).rejects.toThrow(
			'SAVE_SIGNATURES_STRIPPED',
		);
		await expect(saveEditorDocument(handler, signedSnapshot)).rejects.toThrow(
			'SAVE_SIGNATURES_STRIPPED',
		);
	} finally {
		handler.dispose();
	}
});
