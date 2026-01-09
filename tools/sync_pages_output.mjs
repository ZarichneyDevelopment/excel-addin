#!/usr/bin/env node
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const repoRoot = path.resolve(__dirname, '..');
const distDir = path.join(repoRoot, 'budget-helper', 'dist');
const docsDir = path.join(repoRoot, 'docs');

function exists(p) {
  try {
    fs.accessSync(p);
    return true;
  } catch {
    return false;
  }
}

function ensureDir(p) {
  fs.mkdirSync(p, { recursive: true });
}

function rmIfExists(p) {
  if (!exists(p)) return;
  fs.rmSync(p, { recursive: true, force: true });
}

function isGeneratedRootFile(name) {
  if (name === '.nojekyll') return true;
  if (name === 'manifest.xml') return true;
  if (name === 'commands.html' || name === 'commands.js' || name === 'commands.js.map') return true;
  if (name === 'taskpane.html' || name === 'taskpane.js' || name === 'taskpane.js.map' || name === 'taskpane.js.LICENSE.txt') return true;
  if (name === 'polyfill.js' || name === 'polyfill.js.map') return true;
  if (/^[0-9a-f]{12,}\\.css$/i.test(name)) return true;
  return false;
}

function copyFile(src, dest) {
  ensureDir(path.dirname(dest));
  fs.copyFileSync(src, dest);
}

function copyDir(srcDir, destDir) {
  ensureDir(destDir);
  for (const entry of fs.readdirSync(srcDir, { withFileTypes: true })) {
    const src = path.join(srcDir, entry.name);
    const dest = path.join(destDir, entry.name);
    if (entry.isDirectory()) {
      copyDir(src, dest);
    } else if (entry.isFile()) {
      copyFile(src, dest);
    }
  }
}

function main() {
  if (!exists(distDir)) {
    console.error(`Expected build output at ${distDir} (run \`cd budget-helper && npm run build\` first).`);
    process.exit(1);
  }

  ensureDir(docsDir);

  // Clean previously generated artifacts (without touching markdown/docs content).
  for (const name of fs.readdirSync(docsDir)) {
    if (!isGeneratedRootFile(name)) continue;
    rmIfExists(path.join(docsDir, name));
  }

  rmIfExists(path.join(docsDir, 'assets'));

  // Copy current build artifacts into docs/ (GitHub Pages "docs/" source mode).
  for (const entry of fs.readdirSync(distDir, { withFileTypes: true })) {
    const src = path.join(distDir, entry.name);
    const dest = path.join(docsDir, entry.name);
    if (entry.isDirectory()) {
      copyDir(src, dest);
    } else if (entry.isFile()) {
      copyFile(src, dest);
    }
  }

  // Disable Jekyll processing so raw assets are served as-is.
  const noJekyllPath = path.join(docsDir, '.nojekyll');
  if (!exists(noJekyllPath)) {
    fs.writeFileSync(noJekyllPath, '');
  }
}

main();

