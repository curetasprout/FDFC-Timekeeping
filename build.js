// build.js — inline vendor + core + app into a single self-contained index.html.
// Plain Node, zero dependencies. Run: node build.js
import { readFileSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = dirname(fileURLToPath(import.meta.url));
const read = (...p) => readFileSync(join(root, ...p), 'utf-8');

// Strip ES-module syntax so files share one script scope (concatenation order matters).
function deModule(src) {
  return src
    .replace(/^\s*import[^\n]*\n/gm, '')
    .replace(/^\s*export\s*\{[^}]*\};?\s*$/gm, '')
    .replace(/^export\s+(function|const|let|var|class|async)/gm, '$1');
}

const vendor = read('vendor', 'exceljs.min.js');
const core = ['parse', 'shift', 'dateutil', 'mapping', 'validate', 'generate', 'exportXlsx']
  .map((m) => `// ---- core/${m}.js ----\n` + deModule(read('src', 'core', `${m}.js`)))
  .join('\n\n');
const app = deModule(read('src', 'app', 'app.js'));

const html = read('src', 'app', 'index.template.html')
  .replace('/*__VENDOR__*/', () => vendor)
  .replace('/*__CORE__*/', () => core)
  .replace('/*__APP__*/', () => app);

writeFileSync(join(root, 'index.html'), html);
console.log(`Built index.html  (${(html.length / 1024).toFixed(0)} KB)`);
