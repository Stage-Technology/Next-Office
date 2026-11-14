/**
 * Generates src/styles/tokens.css from the brand pack's tokens.json (the source of truth).
 * Run with `npm run tokens` after changing tokens.json, and commit both files.
 */
import { readFileSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const tokens = JSON.parse(readFileSync(join(root, 'src/assets/brand/tokens.json'), 'utf8'));
const out = join(root, 'src/styles/tokens.css');

/** `{brand}` is an alias of another token. */
const value = (v) => v.replace(/^\{([\w-]+)\}$/, 'var(--$1)');
const decl = (name, v) => `  --${name}: ${value(v)};`;

const themed = (theme) => [
  ...tokens.color.tokens.map((t) => decl(t.name, t.value[theme])),
  ...tokens.shadow.tokens.map((t) => decl(t.name, t.value[theme])),
];

const fixed = [
  decl('font-sans', tokens.type.families.sans),
  decl('font-mono', tokens.type.families.mono),
  ...tokens.spacing.tokens.map((t) => decl(t.name, t.value)),
  ...tokens.radius.tokens.map((t) => decl(t.name, t.value)),
  ...tokens.size.tokens.map((t) => decl(t.name, t.value)),
];

const styles = tokens.type.groups.flatMap((g) => g.styles.map((s) => ({ ...s, family: g.family })));
const typeClasses = styles.map((s) =>
  [
    `.t-${s.name} {`,
    `  font-family: var(--font-${s.family});`,
    `  font-size: ${s.fontSize};`,
    `  line-height: ${s.lineHeight};`,
    `  font-weight: ${s.fontWeight};`,
    ...(s.letterSpacing ? [`  letter-spacing: ${s.letterSpacing};`] : []),
    '}',
  ].join('\n'),
);

const fonts = tokens.type.fonts.map((f) =>
  [
    '@font-face {',
    `  font-family: "${f.family}";`,
    `  src: url("../assets/brand/${f.file}") format("woff2");`,
    `  font-weight: ${f.weight};`,
    '  font-style: normal;',
    '  font-display: swap;',
    '}',
  ].join('\n'),
);

const css = `/* Generated from src/assets/brand/tokens.json by scripts/build-tokens.mjs. Do not edit by hand. */

:root,
[data-theme="light"] {
${themed('light').join('\n')}
  color-scheme: light;
}

[data-theme="dark"] {
${themed('dark').join('\n')}
  color-scheme: dark;
}

:root {
${fixed.join('\n')}
}

${fonts.join('\n\n')}

/* Type styles */
${typeClasses.join('\n')}
`;

writeFileSync(out, css);
console.log(`Wrote ${out}`);
