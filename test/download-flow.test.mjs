import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import test from 'node:test';
import vm from 'node:vm';

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
  assert.ok(responseBuilder.includes('window.top.postMessage'), 'response should reach the top-level page');
  assert.equal(responseBuilder.includes('parent.postMessage'), false, 'response should not target the intermediate Apps Script iframe');
  assert.equal(responseBuilder.includes(String.raw`<\\/script>`), false, 'response should not emit an escaped closing tag');
});

test('Apps Script POST does not depend on cross-origin iframe postMessage', () => {
  const postStart = indexHtml.indexOf('function postToGasViaHiddenForm');
  const postSource = indexHtml.slice(postStart).split('</script>')[0];

  assert.ok(postSource.includes("mode: 'no-cors'"), 'POST should work without readable cross-origin response headers');
  assert.ok(postSource.includes('new URLSearchParams'), 'POST should keep the Apps Script form-compatible body');
  assert.ok(postSource.includes('networkOnly: true'), 'network completion should be enough for the legacy deployment');
  assert.doesNotMatch(postSource, /keepalive:\s*true/, 'large Base64 image payloads should not use keepalive quota');
  assert.doesNotMatch(postSource, /createElement\('iframe'\)/, 'POST should not wait for an iframe callback');
});

test('Apps Script POST resolves after the network request accepts a large payload', async () => {
  const postStart = indexHtml.indexOf('async function postToGasViaHiddenForm');
  const postSource = indexHtml.slice(postStart).split('\n// ══════════════════════════════════════════════')[0];
  let request;
  const context = {
    AbortController,
    GAS_URL: 'https://example.test/gas',
    URLSearchParams,
    clearTimeout,
    fetch: async (...args) => {
      request = args;
      return { type: 'opaque' };
    },
    setTimeout
  };

  vm.runInNewContext(`${postSource}\nthis.post = postToGasViaHiddenForm;`, context);
  const result = await context.post({ imageBase64: 'x'.repeat(100_000) });

  assert.equal(result.success, true);
  assert.equal(result.networkOnly, true);
  assert.equal(request[0], 'https://example.test/gas');
  assert.equal(request[1].method, 'POST');
  assert.equal(request[1].mode, 'no-cors');
  assert.ok(request[1].body.get('payload').length > 100_000);
});
