'use strict';
const assert = require('assert');
const fs = require('fs');
const path = require('path');
const os = require('os');
const crypto = require('crypto');
const { retainPayload, extractStream, run } = require(process.argv[2]);
const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'pst-checkpoint-'));
const image = Buffer.from([255, 216, 0, 255, 217]);
const imageHash = crypto.createHash('sha256').update(image).digest('hex');
assert.strictEqual(retainPayload('photo.jpg', image, imageHash, new Set()), false);
assert.strictEqual(retainPayload('photo.jpg', image, imageHash, new Set([imageHash])), true);
assert.strictEqual(retainPayload('wrong-name.bin', Buffer.from('d0cf11e0a1b11ae1', 'hex'), '', new Set()), true);
const payload = path.join(directory, 'not-retained.bin');
let offset = 0;
const stream = { length: image.length, readCompletely: block => { image.copy(block, 0, offset); offset += block.length; } };
const hashed = extractStream(stream, payload, { shouldRetain: () => false });
assert.strictEqual(hashed.sha256, imageHash);
assert.strictEqual(fs.existsSync(payload), false);
const fakeReader = path.join(directory, 'fake-reader.cjs');
fs.writeFileSync(fakeReader, `
const bytes=Buffer.from('sheet\\tdata\\n');
exports.PSTFile=class {
 close(){}
 getRootFolder(){let done=false;return {descriptorNodeId:1,displayName:'root',contentCount:1,hasSubfolders:false,
 getNextChild(){if(done)return null;done=true;return {descriptorNodeId:2,numberOfAttachments:1,subject:'test',
 getAttachment(){let position=0;return {attachMethod:1,longFilename:'source.xls',mimeTag:'application/vnd.ms-excel',
 fileInputStream:{length:bytes.length,readCompletely(block){bytes.copy(block,0,position);position+=block.length;}}};}};}};}
};
`);
const job = path.join(directory, 'job.json');
fs.writeFileSync(job, JSON.stringify({ output: directory, reader_module: fakeReader,
  pst_files: ['fixture.pst'], retain_payloads: 'candidates_and_targets', known_target_sizes: [] }));
run(job);
const journal = fs.readFileSync(path.join(directory, 'attachments.jsonl'));
fs.writeFileSync(job, JSON.stringify({ output: directory, reader_module: fakeReader,
  pst_files: ['fixture.pst'], append_journal: true, retain_payloads: 'candidates_and_targets', known_target_sizes: [] }));
run(job);
assert.deepStrictEqual(fs.readFileSync(path.join(directory, 'attachments.jsonl')), journal);
assert.strictEqual(JSON.parse(fs.readFileSync(path.join(directory, 'reader-status.json'))).counts.messages, 1);
console.log('Retention and restart checks passed. Test artifacts: ' + directory);
