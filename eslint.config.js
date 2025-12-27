import js from '@eslint/js';
import tseslint from 'typescript-eslint';
import prettier from 'eslint-config-prettier';

export default [
  js.configs.recommended,
  ...tseslint.configs.recommended,
  prettier,
  {
    files: ['src/**/*.ts'],
    languageOptions: {
      globals: {
        SpreadsheetApp: 'readonly',
        GoogleAppsScript: 'readonly',
      },
    },
    rules: {
      // Allow unused vars for Google Apps Script entry points (called by Apps Script, not by code)
      '@typescript-eslint/no-unused-vars': [
        'error',
        {
          varsIgnorePattern: '^(onOpen|copyDynamicRange|hapusNilaiDariRentang|setFormulasBatch|setProfitBatch|isiFormulasticker|isiFormulaProfitSticker|processYellowRows)$',
        },
      ],
    },
  },
  {
    ignores: ['build/**', 'node_modules/**'],
  },
];
