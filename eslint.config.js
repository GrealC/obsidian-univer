import antfu from '@antfu/eslint-config'

export default antfu({
  ignores: ['**/.pnpm-store/**', '**/data/**.ts', '**/dist/**'],
}, {
  rules: {
    'no-new': 'off',
    'ts/ban-ts-comment': 'off',
    'unused-imports/no-unused-vars': 'off',
  },
})
