'use strict';
// Read-only PST traversal. Payload filenames are hashes, never sender paths.
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

function retry(action) {
  for (let attempt = 0; ; attempt++) {
    try { return action(); }
    catch (error) {
      if (attempt === 7 || !['EPERM', 'EACCES', 'EBUSY'].includes(error.code)) throw error;
      Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, 100 * (attempt + 1));
    }
  }
}

function saveJson(filename, value) {
  retry(() => {
    fs.writeFileSync(filename + '.tmp', JSON.stringify(value, null, 2) + '\n');
    fs.renameSync(filename + '.tmp', filename);
  });
}

function extractStream(stream, filename, options = {}) {
  const length = Number(stream.length.toString());
  if (!Number.isSafeInteger(length) || length < 0 || length > 2 * 1024 ** 3) {
    throw new Error('Invalid or oversized attachment stream length');
  }
  const hash = crypto.createHash('sha256');
  const md5 = crypto.createHash('md5');
  let fd = null;
  let stored = false;
  let written = 0;
  let prefix = Buffer.alloc(0);
  try {
    while (written < length) {
      const block = Buffer.alloc(Math.min(1024 * 1024, length - written));
      stream.readCompletely(block);
      if (!written) {
        prefix = Buffer.from(block.subarray(0, 4096));
        if (!options.shouldRetain || options.shouldRetain(prefix, length)) {
          fd = fs.openSync(filename, 'w');
          stored = true;
        }
      }
      hash.update(block);
      md5.update(block);
      let offset = 0;
      while (fd !== null && offset < block.length) offset += fs.writeSync(fd, block, offset, block.length - offset);
      written += block.length;
    }
    if (!length && (!options.shouldRetain || options.shouldRetain(prefix, length))) {
      fd = fs.openSync(filename, 'w');
      stored = true;
    }
  } finally { if (fd !== null) fs.closeSync(fd); }
  return { sha256: hash.digest('hex'), md5: md5.digest('hex'), bytes: written, _prefix: prefix, _stored: stored };
}

function retainPayload(filename, prefix, sha256, targetHashes) {
  if (targetHashes.has(sha256)) return true;
  if (/\.(xls\w*|xlt\w*|xlw|xlk|xl|csv|tsv|sylk|slk|wk\d|wks|doc\w*|dot\w*|ppt\w*|pot\w*|rtf|html?|eml|msg)$/i.test(filename || '')) return true;
  if (prefix.subarray(0, 8).equals(Buffer.from('d0cf11e0a1b11ae1', 'hex'))) return true;
  if (prefix.subarray(0, 2).equals(Buffer.from('PK'))) return true;
  if (/^(\{\\rtf|%PDF)/.test(prefix.toString('latin1'))) return true;
  if (prefix.length >= 4 && [0x0009, 0x0209, 0x0409, 0x0809].includes(prefix.readUInt16LE(0))) return true;
  if (prefix.subarray(0, 4).equals(Buffer.from('00000200', 'hex'))) return true;
  if (prefix.length >= 2 && [0xfeff, 0xfffe].includes(prefix.readUInt16LE(0))) return true;
  if (!prefix.length) return false;
  // Keep text-like payloads, including delimiters and legacy text spreadsheets.
  const textBytes = Array.from(prefix).filter(byte => byte === 9 || byte === 10 || byte === 13 || byte >= 32).length;
  return !prefix.includes(0) && textBytes / prefix.length >= 0.85;
}

function run(jobFile) {
  const job = JSON.parse(fs.readFileSync(jobFile, 'utf8').replace(/^\uFEFF/, ''));
  const { PSTFile } = require(job.reader_module);
  const output = job.output;
  fs.mkdirSync(path.join(output, 'attachments'), { recursive: true });
  const journal = path.join(output, 'attachments.jsonl');
  const errorJournal = path.join(output, 'reader-issues.jsonl');
  const checkpointFile = path.join(output, 'reader-checkpoints.json');
  let checkpoint = job.append_journal && fs.existsSync(checkpointFile)
    ? JSON.parse(fs.readFileSync(checkpointFile, 'utf8')) : null;
  if (checkpoint) {
    if (fs.statSync(journal).size < checkpoint.journal_bytes || fs.statSync(errorJournal).size < checkpoint.issue_journal_bytes)
      throw new Error('Reader journals are shorter than their saved checkpoint');
    fs.truncateSync(journal, checkpoint.journal_bytes);
    fs.truncateSync(errorJournal, checkpoint.issue_journal_bytes);
  } else {
    fs.writeFileSync(journal, '');
    fs.writeFileSync(errorJournal, '');
  }
  const counts = Object.assign({ pst_files: 0, folders: 0, messages: 0, attachment_payloads: 0,
    embedded_messages: 0, references_without_payload: 0, issues: 0 }, checkpoint ? checkpoint.counts : {});
  const completed = new Set(checkpoint ? checkpoint.completed_files : []);
  const targetHashes = new Set(job.known_target_sha256 || []);
  const targetSizes = new Set(job.known_target_sizes || []);
  let currentPst = null;
  const progress = () => saveJson(path.join(output, 'reader-status.json'),
    { state: 'running', updated_utc: new Date().toISOString(), current_pst: currentPst, counts });
  function issue(context, error) {
    counts.issues++;
    fs.appendFileSync(errorJournal, JSON.stringify({ ...context, error: String(error.stack || error) }) + '\n');
  }
  function optional(object, field, context) {
    try {
      const value = object[field];
      return value instanceof Date ? (Number.isFinite(value.getTime()) ? value.toISOString() : null) : value;
    } catch (error) { issue({ ...context, field, stage: 'message_metadata' }, error); return null; }
  }
  function message(message, folders, chain, depth = 0) {
    if (depth > 16) throw new Error('Embedded-message nesting limit exceeded');
    const id = message.descriptorNodeId.toString();
    const context = { pst: currentPst, folder_path: folders, message_node_id: id, embedded_message_chain: chain };
    counts.messages++;
    const metadata = { ...context, internet_message_id: optional(message, 'internetMessageId', context),
      subject: optional(message, 'subject', context), sender: optional(message, 'senderEmailAddress', context),
      delivered: optional(message, 'messageDeliveryTime', context) };
    const total = message.numberOfAttachments;
    for (let index = 0; index < total; index++) {
      const attachmentContext = { ...metadata, attachment_index: index };
      try {
        const attachment = message.getAttachment(index);
        const method = attachment.attachMethod;
        const record = { ...attachmentContext, method, original_filename: attachment.longFilename || attachment.filename || null,
          mime_type: attachment.mimeTag || null };
        if (method === 5) {
          const embedded = attachment.embeddedPSTMessage;
          if (!embedded) throw new Error('Embedded-message attachment has no readable message');
          counts.embedded_messages++;
          fs.appendFileSync(journal, JSON.stringify({ ...record, type: 'embedded_message' }) + '\n');
          messageWalk(embedded, folders, [...chain, { message_node_id: id, attachment_index: index }], depth + 1);
          continue;
        }
        const stream = attachment.fileInputStream;
        if (!stream) {
          if ([2, 3, 4, 7].includes(method)) {
            counts.references_without_payload++;
            fs.appendFileSync(journal, JSON.stringify({ ...record, type: 'reference_without_payload' }) + '\n');
            continue;
          }
          throw new Error('Attachment has no readable binary payload');
        }
        const pending = path.join(output, 'attachments', 'pending-' + process.pid + '.bin');
        const options = job.retain_payloads === 'candidates_and_targets'
          ? { shouldRetain: (prefix, size) => targetSizes.has(size)
            || (targetHashes.size > 0 && !job.known_target_sizes)
            || retainPayload(record.original_filename, prefix, '', targetHashes) } : {};
        const { _prefix, _stored, ...hashes } = extractStream(stream, pending, options);
        const retained = job.retain_payloads !== 'candidates_and_targets'
          || retainPayload(record.original_filename, _prefix, hashes.sha256, targetHashes);
        const relative = retained ? 'attachments/' + hashes.sha256 + '.bin' : null;
        if (retained) {
          if (!_stored) throw new Error('Required attachment payload was not staged');
          retry(() => fs.renameSync(pending, path.join(output, relative)));
        } else if (_stored) fs.unlinkSync(pending);
        fs.appendFileSync(journal, JSON.stringify({ ...record, type: 'payload', ...hashes, path: relative, retained }) + '\n');
        counts.attachment_payloads++;
        if (counts.attachment_payloads % 10 === 0) progress();
      } catch (error) { issue({ ...attachmentContext, stage: 'attachment' }, error); }
    }
    if (counts.messages % 50 === 0) progress();
  }
  // Separate name so an embedded message does not shadow the traversal function.
  const messageWalk = message;
  for (const filename of job.pst_files) {
    currentPst = path.basename(filename);
    if (completed.has(currentPst)) continue;
    const visited = new Set();
    let pst;
    function folder(node, parents) {
      const id = node.descriptorNodeId.toString();
      if (visited.has(id)) throw new Error('Repeated PST folder node: ' + id);
      visited.add(id);
      counts.folders++;
      const location = [...parents, node.displayName || '(root)'];
      if (node.contentCount > 0) {
        while (true) {
          let child;
          try { child = node.getNextChild(); }
          catch (error) { issue({ pst: currentPst, folder_path: location, stage: 'folder_messages' }, error); break; }
          if (!child) break;
          try { messageWalk(child, location, []); }
          catch (error) { issue({ pst: currentPst, folder_path: location, stage: 'message' }, error); }
        }
      }
      if (node.hasSubfolders) {
        for (const child of node.getSubFolders()) {
          try { folder(child, location); }
          catch (error) { issue({ pst: currentPst, folder_path: location, stage: 'subfolder' }, error); }
        }
      }
    }
    try {
      pst = new PSTFile(filename);
      folder(pst.getRootFolder(), []);
      counts.pst_files++;
      completed.add(currentPst);
    } catch (error) { issue({ pst: currentPst, stage: 'pst_file' }, error); }
    finally { if (pst) pst.close(); }
    // An attempted PST with issues remains explicitly reported rather than
    // being duplicated in the journal on an unattended restart.
    completed.add(currentPst);
    checkpoint = { completed_files: Array.from(completed), counts,
      journal_bytes: fs.statSync(journal).size, issue_journal_bytes: fs.statSync(errorJournal).size };
    saveJson(checkpointFile, checkpoint);
    progress();
  }
  saveJson(path.join(output, 'reader-status.json'),
    { state: counts.issues ? 'complete_with_issues' : 'complete', counts, finished_utc: new Date().toISOString() });
  if (!counts.pst_files) throw new Error('No PST completed traversal; inspect reader-issues.jsonl');
  console.log(JSON.stringify(counts));
}

module.exports = { extractStream, retainPayload, run };
if (require.main === module) run(process.argv[2]);
