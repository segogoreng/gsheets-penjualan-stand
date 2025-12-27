import resolve from '@rollup/plugin-node-resolve';
import babel from '@rollup/plugin-babel';

const extensions = ['.ts', '.js'];

// Custom plugin to prevent tree-shaking for Google Apps Script
// Apps Script needs all top-level functions to be available
function preventTreeShakingPlugin() {
  return {
    name: 'no-treeshaking',
    resolveId(id, importer) {
      if (!importer) {
        // Entry point - prevent tree-shaking
        return { id, moduleSideEffects: 'no-treeshake' };
      }
      return null;
    },
  };
}

export default {
  input: './src/index.ts',
  output: {
    dir: 'build',
    format: 'cjs',
  },
  plugins: [
    preventTreeShakingPlugin(),
    resolve({
      extensions,
    }),
    babel({
      extensions,
      babelHelpers: 'bundled',
      presets: ['@babel/preset-env', '@babel/preset-typescript'],
    }),
  ],
};
