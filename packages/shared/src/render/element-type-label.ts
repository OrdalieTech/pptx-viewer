/** Translation key for a model element's type. */
export function elementTypeLabelKey(type: string): string {
	const name =
		type === 'text'
			? 'textBox'
			: [
						'shape',
						'connector',
						'image',
						'picture',
						'chart',
						'table',
						'smartArt',
						'media',
						'group',
						'ink',
				  ].includes(type)
				? type
				: 'object';
	return `pptx.elementType.${name}`;
}
