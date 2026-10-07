import { fileURLToPath } from 'node:url';

import { defineConfig } from 'vitest/config';

// Exercise owning source fixes without rebuilding shared workspace packages.
export default defineConfig({
	resolve: {
		alias: [
			{
				find: 'pptx-viewer-core',
				replacement: fileURLToPath(new URL('../../../../core/src/index.ts', import.meta.url)),
			},
			{
				find: 'pptx-viewer-shared/collaboration',
				replacement: fileURLToPath(
					new URL('../../../../shared/src/collaboration/index.ts', import.meta.url),
				),
			},
		],
	},
	test: { include: ['src/__tests__/codec/*.test.ts'] },
});
