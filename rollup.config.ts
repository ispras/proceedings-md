import {nodeResolve} from '@rollup/plugin-node-resolve';
import typescript from "@rollup/plugin-typescript";
import commonjs from "@rollup/plugin-commonjs";
import {RollupOptions} from "rollup";

// @ts-ignore
const production = false

const config: RollupOptions = {
    input: 'src/main.ts',
    output: {
        format: 'cjs',
        sourcemap: !production,
        file: 'src/main.js'
    },
    plugins: [
        nodeResolve({
            browser: false
        }),
        typescript({
            tsconfig: "tsconfig.json",
            sourceMap: !production,
            inlineSources: !production,
        }),
        commonjs()
    ],
    external: [
        /node_modules/
    ]
};

export default config