import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const webDir = resolve(root, 'web');
const outputDir = resolve(webDir, 'apps_script');

const [html, css, javascript] = await Promise.all([
  readFile(resolve(webDir, 'index.html'), 'utf8'),
  readFile(resolve(webDir, 'style.css'), 'utf8'),
  readFile(resolve(webDir, 'app.js'), 'utf8'),
]);

const bootstrappedJavascript = `(() => {
  let attempts = 0;
  const initialize = () => {
    if (!document.getElementById('excel-file')) {
      attempts += 1;
      if (attempts < 100) window.setTimeout(initialize, 50);
      return;
    }
    (() => {
${javascript}
    })();
  };
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initialize, { once: true });
  } else {
    initialize();
  }
})();`;

const inlineHtml = html
  .replace(/  <link rel="stylesheet" href="style\.css">\r?\n/, `  <style>\n${css}\n  </style>\n`)
  .replace(/  <script defer src="https:\/\/cdn\.jsdelivr\.net\/npm\/exceljs@4\.4\.0\/dist\/exceljs\.min\.js"><\/script>\r?\n/, '')
  .replace(/  <script defer src="https:\/\/cdn\.jsdelivr\.net\/npm\/jszip@3\.10\.1\/dist\/jszip\.min\.js"><\/script>\r?\n/, '')
  .replace(/  <script defer src="https:\/\/cdnjs\.cloudflare\.com\/ajax\/libs\/pdf\.js\/3\.11\.174\/pdf\.min\.js"><\/script>\r?\n/, '')
  .replace(/  <script defer src="app\.js"><\/script>\r?\n/, `  <script src="https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js"></script>\n  <script src="https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js"></script>\n  <script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"></script>\n  <script>\n${bootstrappedJavascript}\n  </script>\n`);

if (inlineHtml === html || /(?:href|src)="(?:style\.css|app\.js)"/.test(inlineHtml) || /<script defer /.test(inlineHtml)) {
  throw new Error('Could not replace all web assets while preparing the Apps Script HTML.');
}

await mkdir(outputDir, { recursive: true });
await writeFile(resolve(outputDir, 'Index.html'), inlineHtml, 'utf8');
console.log(`Prepared Google Apps Script files in ${outputDir}`);
