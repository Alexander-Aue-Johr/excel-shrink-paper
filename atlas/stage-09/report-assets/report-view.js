(function () {
'use strict';
/** Formula paths explain measured pixel changes in the selected capture. */
function shouldInspectFormulas(metrics) {
  return Number.isFinite(metrics?.changed_pixels) && metrics.changed_pixels > 0;
}

/** Only retained PNGs inside the study evidence tree are addressable by the atelier. */
function evidenceImageUrl(path) {
  if (typeof path !== 'string' || !path.startsWith('evidence/')) return null;
  const parts = path.split('/');
  if (!path.endsWith('.png') || parts.some(part => !part || part === '.' || part === '..'
      || /[\\:\u0000-\u001f]/.test(part))) return null;
  return '/api/evidence/' + parts.slice(1).map(encodeURIComponent).join('/');
}

/** Preserve the pair and phase: O versus SD differs from O versus rebuilt SD (SDR). */
function atelierRegions(report, differencesOnly = true, imageUrl = evidenceImageUrl) {
  const regions = [];
  for (const [pairIndex, pair] of (report.comparisons || []).entries()) {
    const entries = ['findings', 'charts'].flatMap(kind =>
      (pair[kind] || []).map((finding, findingIndex) => ({ finding, key: `${pairIndex}:${kind}:${findingIndex}` })));
    for (const { finding, key } of entries) {
      const different = ['pixels_differ', 'geometry_differs', 'rendering_differs', 'presence_differs']
        .some(flag => finding[flag] === true);
      if (differencesOnly && !different) continue;
      regions.push({
        ...finding, key, different, left: pair.left, right: pair.right,
        label: `${finding.sheet} / ${finding.cell || finding.chart} / ${pair.left} / ${pair.right}`,
        leftUrl: imageUrl(finding.left_png || finding.left_image),
        rightUrl: imageUrl(finding.right_png || finding.right_image),
      });
    }
  }
  return regions;
}
const number = value => Number(value || 0).toLocaleString('en-US');

/** Descriptive measurements, never a severity threshold that certifies a pass. */
function pixelSummary(metrics) {
  if (!metrics) return 'Pixel measurements pending';
  const text = `${number(metrics.changed_pixels)} changed px · max Δ${number(metrics.max_channel_delta)}/255 · ΣΔ ${number(metrics.total_channel_delta)}`;
  const exterior = metrics.unmatched_pixels
    ? ` · ${number(metrics.unmatched_pixels)} unpaired px (${number(metrics.unmatched_nonwhite_pixels)} non-white)` : '';
  const missing = metrics.unmeasured_regions ? ` · ${number(metrics.unmeasured_regions)} regions unavailable` : '';
  return text + exterior + missing;
}

const make = (tag, className, text) => {
  const element = document.createElement(tag);
  element.className = className;
  if (text != null) element.textContent = text;
  return element;
};
const svgNode = (tag, attributes) => {
  const element = document.createElementNS('http://www.w3.org/2000/svg', tag);
  for (const [key, value] of Object.entries(attributes)) element.setAttribute(key, value);
  return element;
};

/** Keep original images intact. All outlines, masks and enlarged crops are views
 * of retained evidence, never new Excel captures or revised comparison verdicts. */
function createPixelInspector(views, metrics, region) {
  const root = make('section', 'atelier-pixel-inspector');
  const summary = make('p', 'atelier-pixel-summary', pixelSummary(metrics));
  const explanation = make('p', 'atelier-pixel-explanation',
    'Red: changed overlapping RGBA pixels. Amber: content without a matching capture pixel. '
    + 'ΣΔ is the sum of absolute channel differences, not a regression verdict. Click an outlined region or either image to inspect it.');
  if (!metrics.changed_pixels && region.geometry_differs) {
    explanation.textContent = 'No overlapping RGBA pixels changed. The comparison reports different capture geometry. '
      + 'Amber outlines locate unpaired content; this may be a cropped image edge rather than a workbook change.';
  }
  const toggle = make('button', '', 'Pause red pulse');
  toggle.type = 'button';
  toggle.addEventListener('click', () => {
    const paused = root.classList.toggle('pulse-paused');
    for (const view of views) view.stage.classList.toggle('pulse-paused', paused);
    toggle.textContent = paused ? 'Resume red pulse' : 'Pause red pulse';
  });
  const crops = make('div', 'atelier-pixel-crops');
  const more = make('button', '', 'Show more enlarged regions');
  more.type = 'button';
  const reset = make('button', '', 'Show detected regions');
  reset.type = 'button';
  reset.hidden = true;
  const coverage = make('p', 'atelier-pixel-explanation');
  root.append(summary, explanation, toggle, coverage, crops, more, reset);
  let mask = null;
  let visibleCount = 4;
  let focused = null;
  const hotspots = metrics.hotspots || [];

  function cropBox(spot) {
    const x = Math.max(0, spot.x - 8), y = Math.max(0, spot.y - 8);
    return {
      x, y,
      width: Math.max(1, Math.min(128, Math.max(40, spot.width + 16), Math.max(metrics.left_width, metrics.right_width) - x)),
      height: Math.max(1, Math.min(96, Math.max(40, spot.height + 16), Math.max(metrics.left_height, metrics.right_height) - y)),
    };
  }

  function paint(canvas, image, box, scale) {
    canvas.width = box.width * scale; canvas.height = box.height * scale;
    const context = canvas.getContext('2d');
    context.imageSmoothingEnabled = false;
    const width = Math.min(box.width, image.naturalWidth - box.x);
    const height = Math.min(box.height, image.naturalHeight - box.y);
    if (width > 0 && height > 0) context.drawImage(image,
      box.x, box.y, width, height, 0, 0, width * scale, height * scale);
  }

  function renderCrops() {
    crops.replaceChildren();
    const selected = focused ? [focused] : hotspots.slice(0, visibleCount);
    coverage.textContent = `${hotspots.length} of ${metrics.hotspot_count || 0} hotspot boxes retained. `
      + 'The red mask contains every changed overlapping pixel. Checkerboard means outside a retained capture.';
    for (const [index, spot] of selected.entries()) {
      const box = cropBox(spot);
      const scale = Math.max(1, Math.min(8, Math.floor(320 / box.width), Math.floor(220 / box.height)));
      const card = make('article', 'atelier-pixel-crop');
      card.append(make('h4', '', `${spot.unpaired ? 'Unpaired content' : 'Pixel region'} ${index + 1} · x ${spot.x}, y ${spot.y}`));
      const pair = make('div', 'atelier-crop-pair');
      for (const view of views) {
        const figure = make('figure', '');
        const stage = make('div', 'atelier-crop-stage');
        const canvas = make('canvas', '');
        paint(canvas, view.image, box, scale);
        stage.append(canvas);
        if (mask) {
          const overlay = make('canvas', 'atelier-diff-pulse');
          paint(overlay, mask, box, scale);
          stage.append(overlay);
        }
        figure.append(make('figcaption', '', view.label), stage);
        pair.append(figure);
      }
      card.append(pair); crops.append(card);
    }
    more.hidden = Boolean(focused) || visibleCount >= hotspots.length;
    reset.hidden = !focused;
    if (!selected.length) crops.append(make('p', '', 'No changed or unpaired non-white pixels to enlarge. Click an image for manual inspection.'));
  }

  function focus(spot) {
    focused = spot;
    renderCrops();
    const reducedMotion = matchMedia('(prefers-reduced-motion: reduce)').matches;
    root.scrollIntoView({ block: 'nearest', behavior: reducedMotion ? 'instant' : 'smooth' });
  }
  reset.addEventListener('click', () => { focused = null; renderCrops(); });
  more.addEventListener('click', () => { visibleCount += 8; renderCrops(); });
  for (const view of views) {
    const svg = svgNode('svg', { viewBox: `0 0 ${view.image.naturalWidth} ${view.image.naturalHeight}`, class: 'atelier-pixel-overlay' });
    if (metrics.overlay_url) svg.append(svgNode('image', {
      href: metrics.overlay_url, x: 0, y: 0, width: metrics.overlap_width,
      height: metrics.overlap_height, class: 'atelier-diff-pulse',
    }));
    for (const spot of hotspots) {
      const rectangle = svgNode('rect', {
        x: Math.max(0, spot.x - 2), y: Math.max(0, spot.y - 2),
        width: spot.width + 4, height: spot.height + 4,
        class: spot.unpaired ? 'unpaired' : 'changed',
        'vector-effect': 'non-scaling-stroke', tabindex: 0, role: 'button',
        'aria-label': `${spot.unpaired ? 'Unpaired' : 'Changed'} pixels at ${spot.x}, ${spot.y}`,
      });
      rectangle.addEventListener('click', () => focus(spot));
      rectangle.addEventListener('keydown', event => { if (event.key === 'Enter') focus(spot); });
      svg.append(rectangle);
    }
    view.stage.append(svg);
    view.image.addEventListener('click', event => {
      const bounds = view.image.getBoundingClientRect();
      focus({ x: Math.floor((event.clientX - bounds.left) * view.image.naturalWidth / bounds.width),
        y: Math.floor((event.clientY - bounds.top) * view.image.naturalHeight / bounds.height), width: 1, height: 1 });
    });
  }
  if (metrics.overlay_url) {
    const image = new Image();
    image.onload = () => { mask = image; if (root.isConnected) renderCrops(); };
    image.src = metrics.overlay_url;
  }
  renderCrops();
  return root;
}
/** Read-only formula presentation shared by the live and portable viewers. */
const formulaElement = (tag, className, text) => {
  const element = document.createElement(tag);
  element.className = className;
  if (text != null) element.textContent = text;
  return element;
};

/** A dependency observation is not an assertion that all differences are volatile. */
function formulaSummary(diagnosis) {
  return [formulaEvidenceSummary(diagnosis), formulaEvidenceSummary(diagnosis?.source_supplement)]
    .filter(Boolean).join(' · ');
}

function formulaEvidenceSummary(diagnosis) {
  const randomPaths = diagnosis?.volatile_paths || 0;
  const environmentPaths = diagnosis?.environment_paths || 0;
  if (!randomPaths && !environmentPaths) return '';
  const kinds = [];
  if (randomPaths) kinds.push(`${randomPaths} random/time dependency paths`);
  if (environmentPaths) kinds.push(`${environmentPaths} filename/environment dependency paths`);
  return diagnosis.evidence === 'prepared_source_cached_values'
    ? `${kinds.join(' · ')} in source · captured values unavailable`
    : `${kinds.join(' · ')} · other differences remain unclassified`;
}

/** Split only the verified call span; all formula text stays text, never HTML. */
function formulaSegments(formula, span) {
  if (!Array.isArray(span) || span.length !== 2 || !span.every(Number.isInteger)
      || span[0] < 0 || span[1] < span[0] || span[1] > formula.length) return [formula, '', ''];
  return [formula.slice(0, span[0]), formula.slice(span[0], span[1]), formula.slice(span[1])];
}

function formulaCode(node, pulsing) {
  const code = formulaElement('code', 'formula-expression');
  const [before, call, after] = formulaSegments(node.formula || '', node.volatile_span ?? node.environment_span);
  code.append(document.createTextNode(before));
  if (call) code.append(formulaElement('span', pulsing ? 'formula-volatile-call' : 'formula-volatile-call pulse-paused', call));
  code.append(document.createTextNode(after));
  return code;
}

/** Use only the retained diagnosis for this exact state pair and region. */
function retainedFormulaDiagnosis(report, regionKey) {
  const entry = (report.formula_diagnostics?.regions || []).find(entry =>
    `${entry.pair_index}:${entry.kind}:${entry.finding_index}` === regionKey);
  return entry?.diagnosis || {
    evidence: 'unavailable',
    reason: report.formula_diagnostics?.omitted_regions
      ? 'No explanation retained for this region; the recorded diagnostic budget omitted some regions.'
      : 'No formula explanation was retained for this region.',
  };
}

/** Render retained paths without API requests or changes to workbook cells. */
function renderFormulaDiagnosis(root, region, diagnosis) {
  renderFormulaEvidence(root, region, diagnosis);
  if (diagnosis.source_supplement) {
    const supplement = formulaElement('section', 'formula-source-supplement');
    const evidence = formulaElement('div', '');
    supplement.append(formulaElement('h4', '', 'Additional source evidence'));
    renderFormulaEvidence(evidence, region, diagnosis.source_supplement);
    supplement.append(evidence);
    root.append(supplement);
  }
}

function renderFormulaEvidence(root, region, diagnosis) {
  root.replaceChildren();
  const source = diagnosis.evidence === 'prepared_source_cached_values';
  const note = formulaElement('p', 'formula-evidence-note', diagnosis.reason || diagnosis.note || 'No formula evidence retained.');
  root.append(formulaElement('h3', '', 'Formula dependencies'), note);
  root.classList.toggle('formula-source-only', source);
  if (diagnosis.changed_value_cells != null) {
    root.append(formulaElement('p', '', `${diagnosis.changed_value_cells} changed saved cell values · `
      + `${diagnosis.examined_cells} examined · ${diagnosis.volatile_paths || 0} random/time dependency paths`
      + ` · ${diagnosis.environment_paths || 0} filename/environment dependency paths`
      + (diagnosis.coverage_complete ? '' : ' · partial diagnostic coverage')));
  }
  const examples = (diagnosis.examples || []).filter(example => example.status === 'path_found');
  if (!examples.length) {
    root.append(formulaElement('p', '', 'No supported random/time or filename dependency path is available. This does not prove absence of other dependencies.'));
    return;
  }
  const badge = formulaElement('p', 'formula-dependency-badge', formulaSummary(diagnosis));
  const select = formulaElement('select', '');
  select.setAttribute('aria-label', 'Cell dependency path');
  for (const example of examples) {
    const start = example.path[0];
    const end = example.path.at(-1);
    select.append(formulaElement('option', '', `${start.sheet}!${start.cell} → ${end.cell} · ${example.path.length} cells`));
  }
  const toggle = formulaElement('button', '', 'Pause formula pulse');
  toggle.type = 'button';
  const graph = formulaElement('div', 'formula-path');
  let pulsing = true;
  const render = () => {
    graph.replaceChildren();
    for (const [index, node] of examples[select.selectedIndex].path.entries()) {
      if (index) graph.append(formulaElement('span', 'formula-path-arrow', '→'));
      const card = formulaElement('article', 'formula-cell');
      card.append(formulaElement('strong', '', `${node.sheet}!${node.cell}`));
      card.append(formulaElement('p', 'formula-values', source
        ? `Source cached value: ${node.right_value ?? '(no cached value)'}`
        : `${region.left}: ${node.left_value ?? '(blank)'} → ${region.right}: ${node.right_value ?? '(blank)'}`));
      card.append(formulaCode(node, pulsing));
      if (node.shared_anchor) card.append(formulaElement('small', '', `Shared formula anchor: ${node.shared_anchor}`));
      if (!source && node.left_formula != null && node.left_formula !== node.formula) {
        card.append(formulaElement('p', '', `Left-state formula differs: ${node.left_formula}`));
      }
      graph.append(card);
    }
  };
  select.addEventListener('change', render);
  toggle.addEventListener('click', () => {
    pulsing = !pulsing;
    toggle.textContent = pulsing ? 'Pause formula pulse' : 'Resume formula pulse';
    render();
  });
  root.append(badge, select, toggle, graph);
  render();
}

function createRetainedFormulaInspector(region, diagnosis) {
  const root = formulaElement('section', 'formula-inspector');
  renderFormulaDiagnosis(root, region, diagnosis);
  return root;
}
/** A virtual clock whose reads never change playback state. */
class ReplayClock {
  constructor(duration, now = () => performance.now()) {
    if (!Number.isSafeInteger(duration) || duration < 0) throw new Error('Invalid replay duration.');
    this.duration = duration;
    this.now = now;
    this.anchor = 0;
    this.wall = now();
    this.speed = 1;
    this.playing = false;
  }

  position() {
    const elapsed = this.playing ? (this.now() - this.wall) * this.speed : 0;
    return Math.min(this.duration, Math.max(0, this.anchor + elapsed));
  }

  seek(position) {
    if (!Number.isFinite(position)) throw new Error('Invalid replay position.');
    this.anchor = Math.min(this.duration, Math.max(0, position));
    this.wall = this.now();
  }

  play() {
    if (this.position() >= this.duration) this.seek(0);
    else this.seek(this.position());
    this.playing = true;
  }

  pause() {
    this.seek(this.position());
    this.playing = false;
  }

  setSpeed(speed) {
    if (!Number.isFinite(speed) || speed <= 0 || speed > 10000) throw new Error('Invalid replay speed.');
    this.seek(this.position());
    this.speed = speed;
  }
}

function replayInteger(value, label) {
  if (!Number.isSafeInteger(value) || value < 0) throw new Error(`Invalid replay ${label}.`);
  return value;
}

async function readReplayBytes(stream, limit, signal) {
  const reader = stream.getReader();
  const parts = [];
  let length = 0;
  try {
    while (true) {
      signal?.throwIfAborted();
      const { value, done } = await reader.read();
      signal?.throwIfAborted();
      if (done) break;
      length += value.byteLength;
      if (length > limit) throw new Error('Replay data exceeds its read budget.');
      parts.push(value);
    }
  } catch (error) {
    await reader.cancel().catch(() => {});
    throw error;
  } finally {
    reader.releaseLock();
  }
  const bytes = new Uint8Array(length);
  let offset = 0;
  for (const part of parts) { bytes.set(part, offset); offset += part.byteLength; }
  return bytes;
}

/** Validate the fixed format before its index can select a file or allocate a cache. */
function validateReplayManifest(manifest) {
  if (manifest?.version !== 1 || manifest.channel !== 'run-progress'
      || manifest.encoding !== 'gzip-json-checkpoints'
      || manifest.clock !== 'elapsed-monotonic-milliseconds' || manifest.complete !== true
      || !Array.isArray(manifest.chunks)) throw new Error('Replay history is unavailable or unsupported.');
  let sequence = 0;
  let lastAt = 0;
  manifest.chunks.forEach((chunk, index) => {
    const name = `chunk-${String(index).padStart(6, '0')}.json.gz`;
    const count = replayInteger(chunk.events, 'chunk count');
    const first = replayInteger(chunk.first_at_ms, 'chunk start');
    const last = replayInteger(chunk.last_at_ms, 'chunk end');
    const bytes = replayInteger(chunk.compressed_bytes, 'chunk size');
    if (chunk.file !== name || chunk.first_sequence !== sequence || count < 1 || count > 32
        || first < lastAt || last < first || bytes > 513 * 1024
        || !/^[a-f0-9]{64}$/.test(chunk.sha256)) throw new Error('Invalid replay chunk index.');
    sequence = replayInteger(sequence + count, 'event count');
    lastAt = last;
  });
  if (manifest.events !== sequence || manifest.duration_ms !== lastAt) throw new Error('Replay totals differ.');
  return manifest;
}

/** Fetch same-origin chunks on demand, retaining at most two decoded arrays. */
class ReplayReader {
  constructor(manifestUrl, fetcher = (url, options) => fetch(url, options), base = location.href) {
    const page = new URL(base);
    this.url = new URL(manifestUrl, page);
    if (!/^https?:$/.test(page.protocol) || this.url.origin !== page.origin
        || !this.url.pathname.endsWith('/replay/manifest.json') || this.url.search || this.url.hash) {
      throw new Error('Replay requires a same-origin HTTP or HTTPS static site.');
    }
    this.fetcher = fetcher;
    this.cache = new Map();
    this.manifest = null;
  }

  async fetchBytes(url, limit, signal) {
    const response = await this.fetcher(url.pathname, { signal, redirect: 'error', credentials: 'same-origin' });
    if (!response.ok || !response.body) throw new Error(`Replay request failed (${response.status}).`);
    return readReplayBytes(response.body, limit, signal);
  }

  async load(signal) {
    const bytes = await this.fetchBytes(this.url, 8 * 1024 * 1024, signal);
    this.manifest = validateReplayManifest(JSON.parse(new TextDecoder('utf-8', { fatal: true }).decode(bytes)));
    this.cache.clear();
    return this.manifest;
  }

  async chunk(index, signal) {
    signal?.throwIfAborted();
    if (this.cache.has(index)) {
      const value = this.cache.get(index);
      this.cache.delete(index);
      this.cache.set(index, value);
      return value;
    }
    const entry = this.manifest.chunks[index];
    const bytes = await this.fetchBytes(new URL(entry.file, this.url), 513 * 1024, signal);
    const hash = new Uint8Array(await crypto.subtle.digest('SHA-256', bytes));
    const digest = Array.from(hash, value => value.toString(16).padStart(2, '0')).join('');
    if (bytes.byteLength !== entry.compressed_bytes || digest !== entry.sha256) throw new Error('Replay checksum differs.');
    const decoded = await readReplayBytes(new Blob([bytes]).stream().pipeThrough(new DecompressionStream('gzip')),
      512 * 1024 + 64, signal);
    const events = JSON.parse(new TextDecoder('utf-8', { fatal: true }).decode(decoded));
    if (!Array.isArray(events) || events.length !== entry.events) throw new Error('Replay event count differs.');
    let at = entry.first_at_ms;
    events.forEach((event, position) => {
      const nextAt = replayInteger(event.at_ms, 'timestamp');
      if (event.sequence !== entry.first_sequence + position || nextAt < at
          || !event.state || typeof event.state !== 'object' || Array.isArray(event.state)
          || new TextEncoder().encode(JSON.stringify(event)).length > 64 * 1024) {
        throw new Error('Replay event order or state is invalid.');
      }
      at = nextAt;
    });
    if (events[0].at_ms !== entry.first_at_ms || at !== entry.last_at_ms) throw new Error('Replay timestamp bounds differ.');
    signal?.throwIfAborted();
    this.cache.set(index, events);
    if (this.cache.size > 2) this.cache.delete(this.cache.keys().next().value);
    return events;
  }

  async stateAt(position, signal) {
    signal?.throwIfAborted();
    if (!this.manifest || !Number.isFinite(position)) throw new Error('Replay is not ready.');
    const entries = this.manifest.chunks;
    let low = 0;
    let high = entries.length;
    while (low < high) {
      const middle = Math.floor((low + high) / 2);
      if (entries[middle].first_at_ms <= position) low = middle + 1;
      else high = middle;
    }
    if (low === 0) return null;
    const events = await this.chunk(low - 1, signal);
    let selected = null;
    for (const event of events) {
      if (event.at_ms > position) break;
      selected = event;
    }
    return selected;
  }
}

function replayElement(tag, text) {
  const element = document.createElement(tag);
  if (text != null) element.textContent = text;
  return element;
}

function replayTime(milliseconds) {
  const seconds = Math.floor(milliseconds / 1000);
  return [Math.floor(seconds / 3600), Math.floor(seconds / 60) % 60, seconds % 60]
    .map(value => String(value).padStart(2, '0')).join(':');
}

/** Playback changes this historical progress panel only; final result rows stay final. */
function installReportReplay(content, descriptor) {
  const panel = replayElement('section');
  panel.className = 'report-replay report-details';
  panel.append(replayElement('h2', 'Recorded run replay'));
  const message = replayElement('p');
  panel.append(message);
  content.append(panel);
  if (!descriptor?.available || descriptor.manifest !== 'replay/manifest.json') {
    message.textContent = descriptor?.reason || 'Replay history was not recorded for this run.';
    return;
  }
  message.textContent = 'Run-level checkpoints. Detailed process, sheet, capture and heartbeat history is unavailable in this channel.';
  const load = replayElement('button', 'Load recorded replay');
  const controls = replayElement('fieldset');
  controls.className = 'report-replay-controls';
  controls.disabled = true;
  const play = replayElement('button', 'Play');
  const back = replayElement('button', '−10 s');
  const forward = replayElement('button', '+10 s');
  const restart = replayElement('button', 'Restart');
  const speed = replayElement('select');
  speed.setAttribute('aria-label', 'Replay speed');
  for (const value of [0.1, 1, 10, 100, 1000]) {
    const option = replayElement('option', `${value}×`);
    option.value = String(value);
    option.selected = value === 1;
    speed.append(option);
  }
  const seek = replayElement('input');
  seek.type = 'range';
  seek.min = '0';
  seek.step = '1';
  seek.value = '0';
  seek.setAttribute('aria-label', 'Recorded replay time');
  const time = replayElement('output', '00:00:00');
  const observation = replayElement('p', 'Load history to inspect a recorded checkpoint.');
  observation.setAttribute('aria-live', 'polite');
  const progress = replayElement('progress');
  progress.max = 1;
  progress.value = 0;
  progress.setAttribute('aria-label', 'Historical completed workbook pairs');
  controls.append(play, back, forward, restart, speed, seek, time);
  panel.append(load, controls, progress, observation);

  let reader;
  let clock;
  let pending;
  let generation = 0;
  let busy = false;
  let frame;
  let lastFrame = 0;
  let disposed = false;
  let displayedSequence;

  function displayTime() {
    const position = clock.position();
    seek.value = String(position);
    time.textContent = `${replayTime(position)} / ${replayTime(clock.duration)}`;
    play.textContent = clock.playing ? 'Pause' : 'Play';
  }

  async function showAt(position) {
    pending?.abort();
    pending = new AbortController();
    const current = ++generation;
    busy = true;
    try {
      const event = await reader.stateAt(position, pending.signal);
      if (current !== generation) return;
      if (!event) {
        displayedSequence = undefined;
        observation.textContent = 'No checkpoint had been recorded at this time.';
        progress.value = 0;
        return;
      }
      if (displayedSequence === event.sequence) return;
      displayedSequence = event.sequence;
      const state = event.state;
      const total = Number.isSafeInteger(state.total) && state.total >= 0 ? state.total : 0;
      const completed = Number.isSafeInteger(state.completed) && state.completed >= 0 ? state.completed : 0;
      progress.max = Math.max(1, total);
      progress.value = Math.max(0, Math.min(completed, total));
      observation.textContent = [
        `Checkpoint ${replayTime(event.at_ms)}`, `Phase ${state.state || 'unavailable'}`,
        `${completed} / ${total} completed`, `${state.failed ?? '?'} failed`,
        `${state.reused ?? '?'} reused`, state.current, state.error,
      ].filter(value => value != null && value !== '').join(' · ');
    } catch (error) {
      if (current === generation && error.name !== 'AbortError') {
        displayedSequence = undefined;
        clock.pause();
        observation.textContent = `Replay unavailable: ${error.message}`;
        displayTime();
      }
    } finally {
      if (current === generation) busy = false;
    }
  }

  function tick(now) {
    if (disposed || !clock.playing) return;
    displayTime();
    if (clock.playing && now - lastFrame >= 100) {
      lastFrame = now;
      if (clock.position() >= clock.duration) {
        clock.pause();
        void showAt(clock.duration);
        displayTime();
      } else if (!busy) void showAt(clock.position());
    }
    if (!disposed && clock.playing) frame = requestAnimationFrame(tick);
  }

  function jump(position) {
    clock.seek(position);
    displayTime();
    void showAt(clock.position());
  }

  load.addEventListener('click', async () => {
    load.disabled = true;
    pending?.abort();
    pending = new AbortController();
    try {
      reader = new ReplayReader(descriptor.manifest);
      const manifest = await reader.load(pending.signal);
      clock = new ReplayClock(manifest.duration_ms);
      seek.max = String(clock.duration);
      controls.disabled = false;
      load.hidden = true;
      await showAt(0);
      displayTime();
    } catch (error) {
      observation.textContent = `Replay unavailable: ${error.message}`;
    } finally {
      load.disabled = false;
    }
  });
  play.addEventListener('click', () => {
    if (clock.playing) {
      clock.pause();
      cancelAnimationFrame(frame);
    } else {
      clock.play();
      frame = requestAnimationFrame(tick);
    }
    displayTime();
    void showAt(clock.position());
  });
  back.addEventListener('click', () => jump(clock.position() - 10000));
  forward.addEventListener('click', () => jump(clock.position() + 10000));
  restart.addEventListener('click', () => jump(0));
  seek.addEventListener('input', () => jump(Number(seek.value)));
  speed.addEventListener('change', () => clock.setSpeed(Number(speed.value)));
  window.addEventListener('pagehide', () => {
    disposed = true;
    generation++;
    pending?.abort();
    clock?.pause();
    cancelAnimationFrame(frame);
  }, { once: true });
}

/** Only immutable package images are addressable, on file:// or a static host. */
function portableImageUrl(path) {
  return typeof path === 'string' && /^evidence\/[a-f0-9]{64}\.(png|jpg)$/.test(path)
    ? (globalThis.atlasPresentationImage?.(path, document.baseURI) ?? path) : null;
}

/** Keep navigation below the exported report root without interpreting a URL. */
function portablePageUrl(path) {
  if (typeof path !== 'string' || !path.endsWith('/index.html')
      || /[\\:\u0000-\u001f?#]/.test(path)) return null;
  const parts = path.split('/');
  if (parts.some(part => !part || part === '.' || part === '..')) return null;
  return parts.map(encodeURIComponent).join('/');
}

/** A reused observation keeps its previous scientific scope; it is not a new pass. */
function reportRowMatches(row, query, filter) {
  if (!String(row.name || row.id || '').toLocaleLowerCase().includes(query.toLocaleLowerCase())) return false;
  const status = String(row.status || 'unknown');
  if (filter === 'differences') return Number(row.different_regions || 0) > 0;
  if (filter === 'reused') return Boolean(row.reuse);
  if (filter === 'unavailable') return !row.report || /fail|error|incomplete|interrupted/.test(status);
  return true;
}

const reportElement = (tag, className = '', text = null) => {
  const element = document.createElement(tag);
  if (className) element.className = className;
  if (text != null) element.textContent = text;
  return element;
};

function reportLink(href, label) {
  const link = reportElement('a', '', label);
  link.href = href;
  return link;
}

function reportDetails(label, content) {
  const details = reportElement('details', 'report-details');
  details.append(reportElement('summary', '', label));
  if (content instanceof Element) details.append(content);
  else details.append(reportElement('pre', '', JSON.stringify(content, null, 2)));
  return details;
}

function startReportPage() {
  const content = document.getElementById('report-content');
  const links = document.getElementById('report-links');
  try {
    const data = JSON.parse(document.getElementById('report-data').textContent);
    document.getElementById('report-title').textContent = data.title || 'DeltaSieve report';
    content.replaceChildren();
    if (data.kind === 'run') {
      links.append(reportLink('run.json', 'Run metadata'));
      showRunReport(content, data);
    } else if (data.kind === 'comparison') {
      // The back path is generated from the manifest depth, never from workbook text.
      if (typeof data.back === 'string' && /^(\.\.\/)+index\.html$/.test(data.back)) {
        links.append(reportLink(data.back, 'All workbooks'));
      }
      links.append(reportLink('report.json', 'Scientific report'), reportLink('bundle.json', 'Integrity manifest'));
      showComparisonReport(content, data.report);
    } else {
      throw new Error('Unknown report page type');
    }
  } catch (error) {
    content.replaceChildren(reportElement('p', 'report-warning', `Cannot display retained report: ${error.message}`));
  }
}

function showRunReport(content, data) {
  const rows = Array.isArray(data.pairs) ? data.pairs : [];
  if (data.merge) {
    content.append(reportElement('p', 'report-warning',
      `Merged collection · ${data.merge.distinct_run_snapshots} source snapshots · ${rows.length} retained observations · `
      + `${data.merge.potential_conflicts} potential outcome conflicts. Counts do not represent unique corpus coverage.`));
    content.append(reportDetails('Merge provenance and overlap checks', data.merge));
    const sources = reportElement('nav', 'report-controls');
    for (const source of data.merge.sources || []) {
      if (!/^[a-f0-9]{64}$/.test(source.snapshot_id || '')) continue;
      const page = portablePageUrl(`runs/${source.snapshot_id}/index.html`);
      sources.append(reportLink(page, `Source snapshot ${source.snapshot_id.slice(0, 8)}`));
    }
    content.append(sources);
  }
  if (data.publication_review) {
    const review = data.publication_review;
    content.append(reportElement('p', 'report-warning',
      `Publication review required · ${review.metadata_finding_count ?? review.metadata_findings.length} metadata concerns recorded. `
      + 'Review workbook contents, screenshots and distribution rights before sharing this export.'));
    content.append(reportLink('publication-review.json', 'Publication review inventory'));
  }
  content.append(reportElement('p', 'report-summary',
    `${rows.length.toLocaleString('en-US')} workbook pairs · Run status: ${data.status || 'unknown'}`));
  if (data.created_at || data.finished_at) {
    content.append(reportElement('p', 'report-note',
      `Started: ${data.created_at || 'not recorded'} · Finished: ${data.finished_at || 'not recorded'}`));
  }
  content.append(reportElement('p', 'report-note',
    'This view contains retained observations. Reused results follow the selected OOXML reuse policy. '
    + 'Missing outputs, failures, omitted states and unchecked regions do not establish visual equality.'));
  if (data.interrupted_index) {
    content.append(reportElement('p', 'report-warning',
      'This snapshot was recovered from an unfinished run journal. The displayed run status is its last recorded state; no validation is running in this static view.'));
  }
  if (data.error) content.append(reportElement('p', 'report-warning', `Recorded run failure: ${data.error}`));
  if (data.summary) content.append(reportDetails('Recorded run summary', data.summary));
  if (data.selection) content.append(reportDetails('Validation partition and coverage', data.selection));
  installReportReplay(content, data.replay);
  const controls = reportElement('div', 'report-controls');
  const search = reportElement('input');
  search.type = 'search';
  search.placeholder = 'Search workbook names';
  search.setAttribute('aria-label', 'Search workbook names');
  const filter = reportElement('select');
  filter.setAttribute('aria-label', 'Filter workbook outcomes');
  for (const [value, label] of [
    ['all', 'All pairs'], ['differences', 'Recorded differences'],
    ['unavailable', 'Missing or incomplete evidence'], ['reused', 'Reused observations'],
  ]) {
    const option = reportElement('option', '', label);
    option.value = value;
    filter.append(option);
  }
  controls.append(search, filter);
  const wrapper = reportElement('div', 'report-table-wrapper');
  const table = reportElement('table', 'report-table');
  const head = reportElement('thead');
  const header = reportElement('tr');
  for (const text of ['Workbook', 'Outcome', 'Differing observations', 'Provenance']) {
    header.append(reportElement('th', '', text));
  }
  head.append(header);
  const body = reportElement('tbody');
  table.append(head, body);
  wrapper.append(table);
  const pagination = reportElement('nav', 'report-pagination');
  pagination.setAttribute('aria-label', 'Workbook pages');
  const previous = reportElement('button', '', 'Previous');
  const next = reportElement('button', '', 'Next');
  const range = reportElement('span');
  pagination.append(previous, range, next);
  content.append(controls, wrapper, pagination);
  let page = 0;
  const pageSize = 50;

  function renderRows() {
    const selected = rows.filter(row => reportRowMatches(row, search.value, filter.value));
    page = Math.min(page, Math.max(0, Math.ceil(selected.length / pageSize) - 1));
    body.replaceChildren();
    for (const row of selected.slice(page * pageSize, (page + 1) * pageSize)) {
      const entry = reportElement('tr');
      const name = reportElement('td');
      const url = portablePageUrl(row.page);
      name.append(url ? reportLink(url, row.name) : reportElement('span', '', row.name));
      const error = row.error || (Array.isArray(row.errors) && row.errors.length ? row.errors.join('\n') : null);
      if (error) {
        const text = typeof error === 'string' ? error : JSON.stringify(error);
        name.append(reportElement('small', 'report-row-error', text.length > 2000
          ? `${text.slice(0, 2000)}… See run metadata for the complete record.` : text));
      }
      entry.append(name, reportElement('td', '', row.status || 'unknown'),
        reportElement('td', '', Number.isSafeInteger(row.different_regions) ? row.different_regions.toLocaleString('en-US') : 'Not recorded'));
      const provenance = reportElement('td');
      provenance.append(reportElement('span', '', row.reuse ? 'Earlier observation reused' : row.report ? 'Current observation' : 'No retained comparison'));
      if (row.merge_source) provenance.append(reportDetails('Source run identity', row.merge_source));
      if (row.reuse) provenance.append(reportDetails('Reuse decision', row.reuse));
      entry.append(provenance);
      body.append(entry);
    }
    previous.disabled = page === 0;
    next.disabled = (page + 1) * pageSize >= selected.length;
    range.textContent = selected.length
      ? `${page * pageSize + 1}–${Math.min((page + 1) * pageSize, selected.length)} of ${selected.length.toLocaleString('en-US')}`
      : 'No matching pairs';
  }
  previous.addEventListener('click', () => { page--; renderRows(); });
  next.addEventListener('click', () => { page++; renderRows(); });
  for (const input of [search, filter]) {
    input.addEventListener(input === search ? 'input' : 'change', () => { page = 0; renderRows(); });
  }
  renderRows();
}

function showComparisonReport(content, report) {
  const allRegions = atelierRegions(report, false, portableImageUrl);
  const different = allRegions.filter(region => region.different).length;
  content.append(reportElement('p', 'report-summary',
    `Recorded status: ${report.status || 'unknown'} · ${allRegions.length} observed regions/charts · ${different} differing observations`));
  content.append(reportElement('p', 'report-note',
    'Images are retained evidence from the recorded workbook states. O / OR: original before / after rebuild. '
    + 'SD / SDR: processed workbook before / after rebuild. A marked pixel difference is an observation, not by itself a programme regression.'));
  if (report.status !== 'completed_with_limitations' || report.portable_report?.failure) {
    content.append(reportElement('p', 'report-warning',
      'This comparison did not complete successfully. Available partial captures remain inspectable; absent results are unconfirmed.'));
    if (report.portable_report?.failure) content.append(reportDetails('Recorded failure', report.portable_report.failure));
  }
  const controls = reportElement('div', 'report-controls');
  const previous = reportElement('button', '', 'Previous region');
  const next = reportElement('button', '', 'Next region');
  const select = reportElement('select', 'report-region-select');
  select.setAttribute('aria-label', 'Observed region, chart and comparison phase');
  const fitButton = reportElement('button', '', 'Show 1:1');
  const include = reportElement('input');
  include.type = 'checkbox';
  const includeLabel = reportElement('label');
  includeLabel.append(include, document.createTextNode('Include checked regions without differences'));
  controls.append(previous, next, select, fitButton, includeLabel);
  const note = reportElement('p', 'report-verdict');
  const area = reportElement('section', 'report-capture-area');
  const imageScroll = reportElement('div', 'report-image-scroll');
  const diagnostic = reportElement('div');
  const formulas = reportElement('div');
  area.append(imageScroll, diagnostic, formulas);
  content.append(controls, note, area);
  let regions = [];
  let generation = 0;
  let fit = true;

  function fitImages() {
    const images = [...imageScroll.querySelectorAll('img.atelier-capture')];
    const width = images.reduce((sum, image) => sum + image.naturalWidth, 0);
    const scale = fit && width > 0 ? Math.min(1, imageScroll.clientWidth / width) : 1;
    for (const image of images) {
      if (image.naturalWidth) image.style.width = `${image.naturalWidth * scale}px`;
    }
  }

  function renderRegion() {
    const current = ++generation;
    imageScroll.replaceChildren();
    diagnostic.replaceChildren();
    formulas.replaceChildren();
    const index = select.selectedIndex;
    const region = regions[index];
    previous.disabled = index <= 0;
    next.disabled = index < 0 || index >= regions.length - 1;
    if (!region) {
      note.textContent = allRegions.length
        ? 'No differences were recorded in the observed regions. Include checked regions to inspect their images. This is not a whole-workbook identity claim.'
        : 'No rendered comparison regions are recorded. Consult the state decisions and coverage; omitted images are not evidence of pixel equality.';
      return;
    }
    const observations = [
      region.pixels_differ && 'pixel difference', region.geometry_differs && 'geometry difference',
      region.rendering_differs && 'rendering change', region.presence_differs && 'chart presence change',
      region.blank_margin_ignored && 'blank trailing margin ignored',
      region.capture_extent_ignored && 'comparison restricted to overlapping pixels',
      region.empty_geometry_ignored && 'both regions have no visible area; measured geometry retained',
      region.pixels_compared === false && 'no pixel comparison',
    ].filter(Boolean);
    note.textContent = `${index + 1} / ${regions.length} · ${observations.join(' · ') || 'No difference recorded'}`;
    const pair = reportElement('div', 'atelier-pair');
    const views = [];
    for (const side of ['left', 'right']) {
      const figure = reportElement('figure', 'atelier-figure');
      figure.append(reportElement('figcaption', '', laneLabel(region[side])));
      const url = region[`${side}Url`];
      if (url) {
        const image = reportElement('img', 'atelier-capture');
        image.alt = `${laneLabel(region[side])} · ${region.sheet} · ${region.cell || region.chart}`;
        image.addEventListener('load', fitImages);
        image.addEventListener('error', () => {
          if (current !== generation) return;
          image.replaceWith(reportElement('p', 'report-warning', 'The retained image is unavailable. No replacement was generated.'));
        }, { once: true });
        image.src = url;
        const stage = reportElement('div', 'atelier-image-stage');
        stage.append(image);
        figure.append(stage);
        views.push({ image, stage, label: laneLabel(region[side]) });
      } else {
        figure.append(reportElement('p', '',
          `No retained image for this state (${region[`${side}_rendering`] || 'see presence and geometry evidence'}).`));
      }
      pair.append(figure);
    }
    imageScroll.append(pair);
    imageScroll.scrollTop = imageScroll.scrollLeft = 0;
    fitImages();
    const retained = report.portable_report?.diagnostics?.[region.key];
    if (shouldInspectFormulas(retained?.metrics)) {
      formulas.append(createRetainedFormulaInspector(region, retainedFormulaDiagnosis(report, region.key)));
    }
    if (views.length === 2 && retained?.metrics) {
      diagnostic.append(reportElement('p', 'report-note', 'Loading retained pixel details…'));
      Promise.all(views.map(view => view.image.decode())).then(() => {
        if (current !== generation) return;
        const metrics = { ...retained.metrics, overlay_url: portableImageUrl(retained.metrics.overlay_url) };
        diagnostic.replaceChildren(createPixelInspector(views, metrics, region));
      }).catch(error => {
        if (current === generation) diagnostic.replaceChildren(reportElement('p', 'report-warning',
          `Retained pixel details unavailable: ${error.message}`));
      });
    } else if (retained?.unavailable) {
      diagnostic.append(reportElement('p', 'report-note', retained.unavailable));
    }
  }

  function populateRegions() {
    const key = regions[select.selectedIndex]?.key;
    regions = atelierRegions(report, !include.checked, portableImageUrl);
    select.replaceChildren(...regions.map(region => reportElement('option', '', region.label)));
    const index = regions.findIndex(region => region.key === key);
    if (index >= 0) select.selectedIndex = index;
    renderRegion();
  }
  previous.addEventListener('click', () => { select.selectedIndex--; renderRegion(); });
  next.addEventListener('click', () => { select.selectedIndex++; renderRegion(); });
  select.addEventListener('change', renderRegion);
  include.addEventListener('change', populateRegions);
  fitButton.addEventListener('click', () => { fit = !fit; fitButton.textContent = fit ? 'Show 1:1' : 'Fit images'; fitImages(); });
  new ResizeObserver(fitImages).observe(imageScroll);
  populateRegions();
  appendScientificContext(content, report);
}

function laneLabel(lane) {
  return ({ O: 'Original · before rebuild', OR: 'Original · rebuilt',
    SD: 'Processed · before rebuild', SDR: 'Processed · rebuilt',
    CP: 'Check Performance · before rebuild', CPR: 'Check Performance · rebuilt' })[lane] || lane;
}

function appendScientificContext(content, report) {
  const limitations = reportElement('ul');
  for (const limitation of report.limitations || []) limitations.append(reportElement('li', '', limitation));
  content.append(reportDetails('Coverage and limitations', limitations));
  content.append(reportDetails('Recorded policies and state decisions', {
    capture_policy: report.capture_policy, execution_strategy: report.execution_strategy,
    capture_partition: report.capture_partition, input_identity: report.input_identity,
    state_decisions: report.state_decisions, content_coverage: report.content_coverage,
  }));
  for (const [field, title] of [
    ['height_checks', 'Row-height observations'], ['conditional_formatting', 'Conditional-formatting decisions'],
    ['formula_diagnostics', 'Recorded formula diagnostics'], ['transitions', 'Recorded XML transitions'],
  ]) {
    if (report[field] != null) content.append(reportDetails(title, report[field]));
  }
}

if (typeof document !== 'undefined' && document.getElementById('report-data')) startReportPage();
}());
