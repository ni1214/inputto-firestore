import { existsSync, readFileSync, readdirSync, statSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const docsDir = resolve(root, 'docs');
const indexPath = resolve(docsDir, 'index.html');
const assetsDir = resolve(docsDir, 'assets');

function fail(message) {
  console.error(`Pages build check failed: ${message}`);
  process.exit(1);
}

if (!existsSync(indexPath)) {
  fail('docs/index.html was not generated.');
}

if (!existsSync(assetsDir)) {
  fail('docs/assets was not generated.');
}

const html = readFileSync(indexPath, 'utf8');
const assetRefs = Array.from(html.matchAll(/(?:src|href)="\.\/(assets\/[^"]+)"/g), (match) => match[1]);
const builtAssets = readdirSync(assetsDir, { withFileTypes: true })
  .filter((entry) => entry.isFile() && /^index-[\w-]+\.(js|css)$/.test(entry.name))
  .map((entry) => {
    const path = resolve(docsDir, 'assets', entry.name);
    return {
      name: entry.name,
      path,
      ext: entry.name.endsWith('.js') ? '.js' : '.css',
      mtimeMs: statSync(path).mtimeMs
    };
  });

if (!assetRefs.some((ref) => ref.endsWith('.js'))) {
  fail('docs/index.html does not reference a built JavaScript asset.');
}

if (!assetRefs.some((ref) => ref.endsWith('.css'))) {
  fail('docs/index.html does not reference a built CSS asset.');
}

for (const ref of assetRefs) {
  const assetPath = resolve(docsDir, ref);
  if (!existsSync(assetPath)) {
    fail(`referenced asset is missing: ${ref}`);
  }
  if (statSync(assetPath).size <= 0) {
    fail(`referenced asset is empty: ${ref}`);
  }
}

for (const ext of ['.js', '.css']) {
  const newest = builtAssets
    .filter((asset) => asset.ext === ext)
    .sort((left, right) => right.mtimeMs - left.mtimeMs)[0];
  const referenced = assetRefs
    .filter((ref) => ref.endsWith(ext))
    .map((ref) => {
      const path = resolve(docsDir, ref);
      return {
        ref,
        mtimeMs: statSync(path).mtimeMs
      };
    })[0];

  if (newest && referenced && referenced.mtimeMs < newest.mtimeMs - 1000) {
    fail(`docs/index.html references stale ${ext} asset: ${referenced.ref}; newest is assets/${newest.name}`);
  }
}

if (!html.includes('__inputtoMarkReady')) {
  fail('startup recovery guard is missing from docs/index.html.');
}

console.log(`Pages build check passed: ${assetRefs.length} referenced assets verified.`);
