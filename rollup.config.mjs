import typescript from '@rollup/plugin-typescript';
import terser from '@rollup/plugin-terser';

const output = {
	banner: `/*
 * @license
 * docx-preview <https://github.com/VolodymyrBaydalka/docxjs>
 * Released under Apache License 2.0  <https://github.com/VolodymyrBaydalka/docxjs/blob/master/LICENSE>
 * Copyright Volodymyr Baydalka
 */`,
	sourcemap: true,
}

const umdOutput = {
	...output,
	name: "docx",
	file: 'dist/docx-preview.js',
	format: 'umd',
	globals: {
		jszip: 'JSZip'
	},
};

export default args => {
	const config = {
		input: 'src/docx-preview.ts',
		output: [umdOutput],
		// `inlineSources: true` makes TypeScript embed source text in the emitted
		// sourcemap (`sourcesContent`) rather than just listing paths. Without it,
		// downstream bundlers (webpack, rollup, parcel) fail to resolve sources on
		// consumer machines because the `../src/*.ts` paths don't exist there.
		// See issue #183. Applies to every output (UMD dev/min, ES dev/min) since
		// the TypeScript plugin runs once at the input stage.
		plugins: [typescript({ sourceMap: true, inlineSources: true })]
	}

	if (args.environment == 'BUILD:production')
		config.output = [umdOutput,
			{
				...umdOutput,
				file: 'dist/docx-preview.min.js',
				plugins: [terser()]
			},
			{
				...output,
				file: 'dist/docx-preview.mjs',
				format: 'es',
			},
			{
				...output,
				file: 'dist/docx-preview.min.mjs',
				format: 'es',
				plugins: [terser()]
			}];

	return config
};