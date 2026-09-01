import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import test from 'node:test';

const repoDir = join(dirname(fileURLToPath(import.meta.url)), '..');
const indexHtml = await readFile(join(repoDir, 'index.html'), 'utf8');
const fallbackHtml = await readFile(join(repoDir, '404.html'), 'utf8');
const codeGs = await readFile(join(repoDir, 'Code.gs'), 'utf8');

test('Save prepares the file picker before waiting for Google Apps Script', () => {
  const saveStart = indexHtml.indexOf('async function saveToSheet()');
  const pickerCall = indexHtml.indexOf('await chooseImageSave()', saveStart);
  const gasCall = indexHtml.indexOf('await postToGasViaHiddenForm(payload)', saveStart);

  assert.ok(saveStart >= 0, 'saveToSheet should exist');
  assert.ok(pickerCall > saveStart, 'Save should prepare a local save target');
  assert.ok(pickerCall < gasCall, 'The picker must run before the async server save');
});

test('Save writes or downloads the image even when the server reports no changes', () => {
  const noChangeStart = indexHtml.indexOf('} else if (res && res.noChange)');
  const successStart = indexHtml.indexOf('} else {', noChangeStart);
  const noChangeHandling = indexHtml.slice(noChangeStart, successStart);

  assert.ok(noChangeStart >= 0, 'noChange handling should exist');
  assert.match(noChangeHandling, /saveCanvasToFile|downloadImg/);
});

test('The native picker and delayed-download regression are covered in the page source', () => {
  assert.match(indexHtml, /showSaveFilePicker/);
  assert.match(indexHtml, /suggestedName/);
  assert.doesNotMatch(indexHtml, /setTimeout\(\(\) => downloadImg\(\), 800\)/);
});

test('GitHub Pages fallback stays synchronized with the main page', () => {
  assert.equal(fallbackHtml, indexHtml);
});

test('Apps Script POST response closes its script tag so postMessage can run', () => {
  const responseBuilderStart = codeGs.indexOf('function outputPostResult_');
  const responseBuilder = codeGs.slice(responseBuilderStart, codeGs.indexOf('\nfunction escapeHtml_', responseBuilderStart));

  assert.ok(responseBuilderStart >= 0, 'outputPostResult_ should exist');
  assert.ok(responseBuilder.includes("const scriptClose = '<' + '/script>';"), 'response should build a real closing script tag at runtime');
  assert.ok(responseBuilder.includes('${scriptClose}</body></html>'), 'response should append the runtime closing script tag');
  assert.equal(responseBuilder.includes(String.raw`<\\/script>`), false, 'response should not emit an escaped closing tag');
});
