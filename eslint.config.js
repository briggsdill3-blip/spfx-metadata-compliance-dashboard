const spfxProfile = require('@microsoft/eslint-config-spfx/lib/flat-profiles/react');
const sonarjs = require('eslint-plugin-sonarjs');

module.exports = [
  ...spfxProfile,
  sonarjs.configs.recommended,
  {
    files: ['**/*.ts', '**/*.tsx'],
    languageOptions: {
      parserOptions: {
        tsconfigRootDir: __dirname,
        project: './tsconfig.json'
      }
    }
  }
];
