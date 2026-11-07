'use strict';
const assert = require('assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const crypto = require('crypto');
const { extractStream } = require(process.argv[2]);
const folder = fs.mkdtempSync(path.join(os.tmpdir(), 'pst-reader-test-'));
try {
  const payload = Buffer.from('unchanged attachment bytes');
  let position = 0;
  const stream = { length: { toString: () => String(payload.length) },
    readCompletely: output => { payload.copy(output, 0, position, position + output.length); position += output.length; } };
  const filename = path.join(folder, 'payload.bin');
  const result = extractStream(stream, filename);
  assert.deepStrictEqual(fs.readFileSync(filename), payload);
  assert.strictEqual(result.sha256, crypto.createHash('sha256').update(payload).digest('hex'));
  assert.throws(() => extractStream({ length: 100, readCompletely: () => { throw new Error('unexpected EOF'); } }, filename), /unexpected EOF/);
  assert.throws(() => extractStream({ length: -1 }, filename), /Invalid/);
  console.log('Attachment bytes, checksum and truncated-stream checks passed.');
} finally {
  // Only the two known files in this newly created test directory are removed.
  const filename = path.join(folder, 'payload.bin');
  if (fs.existsSync(filename)) fs.unlinkSync(filename);
  fs.rmdirSync(folder);
}
