'use strict';

/* ---------- CSV schema ---------- */
// Fixed column order written to disk. row_index is always recomputed on
// save (1..N over loaded+session rows combined) so it's not stored on
// row objects; the other 10 fields are.
const CSV_HEADER = [
  'row_index', 'video_id', 'repetition_number', 'exercise_id', 'person',
  'first_frame', 'last_frame', 'front_cam_orientation', 'correctness_score',
  'error_description', 'video_front_problem'
];
const DELIM = ';';

/* ---------- CSV parse/serialize ---------- */

// Full-text parser (not line-based) so quoted fields may contain the
// delimiter, quotes ("" escaping) or embedded newlines, per legacy files.
function parseCSV(text, delimiter) {
  const rows = [];
  let row = [];
  let cur = '';
  let inQuotes = false;
  let i = 0;
  const len = text.length;
  while (i < len) {
    const ch = text[i];
    if (inQuotes) {
      if (ch === '"') {
        if (text[i + 1] === '"') { cur += '"'; i += 2; continue; }
        inQuotes = false; i++; continue;
      }
      cur += ch; i++; continue;
    }
    if (ch === '"') { inQuotes = true; i++; continue; }
    if (ch === delimiter) { row.push(cur); cur = ''; i++; continue; }
    if (ch === '\r') { i++; continue; }
    if (ch === '\n') { row.push(cur); cur = ''; rows.push(row); row = []; i++; continue; }
    cur += ch; i++;
  }
  if (cur.length > 0 || row.length > 0) { row.push(cur); rows.push(row); }
  return rows.filter(r => !(r.length === 1 && r[0] === ''));
}

function csvField(value) {
  const s = (value === undefined || value === null) ? '' : String(value);
  if (s.includes(DELIM) || s.includes('"') || s.includes('\n') || s.includes('\r')) {
    return '"' + s.replace(/"/g, '""') + '"';
  }
  return s;
}

function buildCsvText(allRows) {
  const lines = [CSV_HEADER.map(csvField).join(DELIM)];
  allRows.forEach((r, idx) => {
    lines.push([
      idx + 1, r.video_id, r.repetition_number, r.exercise_id, r.person,
      r.first_frame, r.last_frame, r.front_cam_orientation, r.correctness_score,
      r.error_description || '', r.video_front_problem || ''
    ].map(csvField).join(DELIM));
  });
  // UTF-8 BOM so Excel opens diacritics correctly; CRLF per RFC4180.
  return '﻿' + lines.join('\r\n') + '\r\n';
}

function parseLoadedCsv(text) {
  const rows = parseCSV(text, DELIM);
  if (rows.length === 0) return [];
  return rows.slice(1).map(f => ({
    video_id: f[1] || '',
    repetition_number: f[2] || '',
    exercise_id: f[3] || '',
    person: f[4] || '',
    first_frame: f[5] || '',
    last_frame: f[6] || '',
    front_cam_orientation: f[7] || '',
    correctness_score: f[8] || '',
    error_description: f[9] || '',
    video_front_problem: f[10] || ''
  }));
}

/* ---------- State ---------- */

let loadedRows = [];      // rows parsed from an appended existing CSV
let sessionRows = [];     // rows added this session
let csvFileHandle = null; // File System Access handle, if available
let videoLoaded = false;
let pendingStart = null;
let pendingEnd = null;
let markPhase = 'start'; // 'start' | 'end' — drives the single toggle button
let seekMainDragging = false;
let zoomDragging = false;

/* ---------- Element refs ---------- */

const video = document.getElementById('video');
const btnOpenVideo = document.getElementById('btnOpenVideo');
const videoFileInput = document.getElementById('videoFileInput');
const videoFileName = document.getElementById('videoFileName');
const videoIdInput = document.getElementById('videoIdInput');
const fpsInput = document.getElementById('fpsInput');
const personInput = document.getElementById('personInput');
const btnNewPerson = document.getElementById('btnNewPerson');

const btnLoadCsv = document.getElementById('btnLoadCsv');
const csvFileInputFallback = document.getElementById('csvFileInputFallback');
const csvFileNameEl = document.getElementById('csvFileName');
const btnForgetCsv = document.getElementById('btnForgetCsv');

const frameDisplay = document.getElementById('frameDisplay');
const seekMain = document.getElementById('seekMain');
const timelineMarks = document.getElementById('timelineMarks');
const seekZoom = document.getElementById('seekZoom');
const zoomWindowSize = document.getElementById('zoomWindowSize');

const btnPlayPause = document.getElementById('btnPlayPause');
const btnStepBack = document.getElementById('btnStepBack');
const btnStepFwd = document.getElementById('btnStepFwd');
const btnJumpBack = document.getElementById('btnJumpBack');
const btnJumpFwd = document.getElementById('btnJumpFwd');
const speedSelect = document.getElementById('speedSelect');

const btnMarkToggle = document.getElementById('btnMarkToggle');
const btnClearMarks = document.getElementById('btnClearMarks');
const startDisplay = document.getElementById('startDisplay');
const endDisplay = document.getElementById('endDisplay');

const repNumberInput = document.getElementById('repNumberInput');
const exerciseIdInput = document.getElementById('exerciseIdInput');
const orientationSelect = document.getElementById('orientationSelect');
const orientationOther = document.getElementById('orientationOther');

const repTableBody = document.getElementById('repTableBody');
const btnClearSession = document.getElementById('btnClearSession');
const overlapWarning = document.getElementById('overlapWarning');
const btnSave = document.getElementById('btnSave');
const statusBar = document.getElementById('statusBar');

const reviewPanel = document.getElementById('reviewPanel');
const reviewTitle = document.getElementById('reviewTitle');
const btnReviewClose = document.getElementById('btnReviewClose');
const btnReviewPlay = document.getElementById('btnReviewPlay');
const reviewStartInput = document.getElementById('reviewStartInput');
const reviewEndInput = document.getElementById('reviewEndInput');
const btnReviewStartMinus = document.getElementById('btnReviewStartMinus');
const btnReviewStartPlus = document.getElementById('btnReviewStartPlus');
const btnReviewEndMinus = document.getElementById('btnReviewEndMinus');
const btnReviewEndPlus = document.getElementById('btnReviewEndPlus');
const btnReviewNext = document.getElementById('btnReviewNext');

/* ---------- Helpers ---------- */

function status(msg, kind) {
  statusBar.textContent = msg;
  statusBar.className = 'status' + (kind ? ' ' + kind : '');
}

function getFps() {
  return parseFloat(fpsInput.value) || 30;
}

function frameFromTime(t) {
  return Math.round(t * getFps());
}

function timeFromFrame(f) {
  return f / getFps();
}

function currentOrientation() {
  return orientationSelect.value === 'other'
    ? orientationOther.value.trim()
    : orientationSelect.value;
}

function nextRepNumber(videoId) {
  let max = 0;
  for (const r of loadedRows.concat(sessionRows)) {
    if (r.video_id === videoId) {
      const n = parseInt(r.repetition_number, 10);
      if (!isNaN(n) && n > max) max = n;
    }
  }
  return max + 1;
}

function updateRepNumberDefault() {
  const id = videoIdInput.value.trim();
  if (id) repNumberInput.value = nextRepNumber(id);
  renderTimelineMarks();
}

// Rows (any mix of loadedRows/sessionRows objects) whose frame range
// overlaps another row of the *same* video_id — compared pairwise after
// sorting by first_frame within each video. Returns a Set of the actual
// row objects involved, usable directly as `.has(row)`.
function computeOverlappingRows(rows) {
  const overlapping = new Set();
  const byVideo = new Map();
  rows.forEach(r => {
    if (!byVideo.has(r.video_id)) byVideo.set(r.video_id, []);
    byVideo.get(r.video_id).push(r);
  });
  byVideo.forEach(list => {
    const sorted = list.slice().sort((a, b) => (parseInt(a.first_frame, 10) || 0) - (parseInt(b.first_frame, 10) || 0));
    for (let i = 0; i < sorted.length - 1; i++) {
      const a = sorted[i], b = sorted[i + 1];
      const aEnd = parseInt(a.last_frame, 10), bStart = parseInt(b.first_frame, 10);
      if (!isNaN(aEnd) && !isNaN(bStart) && aEnd >= bStart) {
        overlapping.add(a);
        overlapping.add(b);
      }
    }
  });
  return overlapping;
}

// Draws every already-saved rep for the current video as a green segment
// (a band spanning its frames plus edge ticks for precision) — red instead
// if it overlaps another rep for this video — plus a blue band and
// green/red ticks for the Start/End mark currently in progress.
function renderTimelineMarks() {
  if (!timelineMarks) return;
  if (!video.duration) { timelineMarks.innerHTML = ''; return; }
  const dur = video.duration;
  const pct = (t) => Math.min(100, Math.max(0, (t / dur) * 100));
  const videoId = videoIdInput.value.trim();
  let html = '';

  const existingForVideo = loadedRows.concat(sessionRows).filter(r => r.video_id === videoId);
  const overlaps = computeOverlappingRows(existingForVideo);
  existingForVideo.forEach(r => {
    const sf = parseInt(r.first_frame, 10);
    const ef = parseInt(r.last_frame, 10);
    const suffix = overlaps.has(r) ? ' existing-overlap' : ' existing';
    if (!isNaN(sf) && !isNaN(ef)) {
      const lo = pct(timeFromFrame(sf));
      const hi = pct(timeFromFrame(ef));
      html += `<div class="mark-range${suffix}" style="left:${lo}%; width:${Math.max(0, hi - lo)}%"></div>`;
    }
    if (!isNaN(sf)) html += `<div class="mark-tick${suffix}" style="left:${pct(timeFromFrame(sf))}%"></div>`;
    if (!isNaN(ef)) html += `<div class="mark-tick${suffix}" style="left:${pct(timeFromFrame(ef))}%"></div>`;
  });

  if (pendingStart !== null && pendingEnd !== null) {
    const lo = pct(timeFromFrame(pendingStart));
    const hi = pct(timeFromFrame(pendingEnd));
    html += `<div class="mark-range pending" style="left:${lo}%; width:${Math.max(0, hi - lo)}%"></div>`;
  }
  if (pendingStart !== null) html += `<div class="mark-tick start" style="left:${pct(timeFromFrame(pendingStart))}%"></div>`;
  if (pendingEnd !== null) html += `<div class="mark-tick end" style="left:${pct(timeFromFrame(pendingEnd))}%"></div>`;

  timelineMarks.innerHTML = html;
}

function setControlsEnabled(enabled) {
  [btnPlayPause, btnStepBack, btnStepFwd, btnJumpBack, btnJumpFwd, speedSelect,
    btnMarkToggle, btnClearMarks].forEach(el => { el.disabled = !enabled; });
}

/* ---------- Video loading ---------- */

btnOpenVideo.addEventListener('click', () => videoFileInput.click());

videoFileInput.addEventListener('change', () => {
  const file = videoFileInput.files[0];
  if (!file) return;
  const url = URL.createObjectURL(file);
  video.src = url;
  video.load();
  videoFileName.textContent = file.name;
  videoIdInput.value = file.name.replace(/\.[^/.]+$/, '');
  zoomInitialized = false; // force the zoom window to recenter for this video's duration
  closeReview();
  // Person is sticky across videos on purpose (a batch is usually one
  // person); use "New Person" if it actually changes.
  resetMarks();
  status(`Loaded video: ${file.name}`);
});

video.addEventListener('loadedmetadata', () => {
  videoLoaded = true;
  seekMain.max = video.duration;
  seekMain.step = 1 / getFps();
  setControlsEnabled(true);
  updateFrameDisplay();
  updateRepNumberDefault();
});

videoIdInput.addEventListener('input', updateRepNumberDefault);
videoIdInput.addEventListener('change', trySaveRepetition);

/* ---------- CSV loading (append mode) ---------- */

btnLoadCsv.addEventListener('click', async () => {
  if (window.showOpenFilePicker) {
    try {
      const [handle] = await window.showOpenFilePicker({
        types: [{ description: 'CSV files', accept: { 'text/csv': ['.csv'] } }]
      });
      csvFileHandle = handle;
      const file = await handle.getFile();
      const text = await file.text();
      onCsvLoaded(text, file.name);
    } catch (err) {
      if (err.name !== 'AbortError') status('Could not open file: ' + err.message, 'error');
    }
  } else {
    csvFileInputFallback.click();
  }
});

csvFileInputFallback.addEventListener('change', () => {
  const file = csvFileInputFallback.files[0];
  if (!file) return;
  csvFileHandle = null; // no write handle in this path; Save will download
  const reader = new FileReader();
  reader.onload = () => onCsvLoaded(reader.result, file.name);
  reader.onerror = () => status('Could not read file.', 'error');
  reader.readAsText(file, 'UTF-8');
});

function onCsvLoaded(text, filename) {
  loadedRows = parseLoadedCsv(text);
  csvFileNameEl.textContent = `${filename} (${loadedRows.length} existing rows)`;
  csvFileNameEl.dataset.rawName = filename;
  btnForgetCsv.disabled = false;
  updateRepNumberDefault();
  status(`Loaded ${loadedRows.length} existing rows from ${filename}. New rows will be appended on Save.`);
}

// Detaches from the loaded file entirely (separate from "Clear All Reps",
// which only touches rows added this session). The file on disk is not
// touched — this just stops treating it as the append target and drops
// its rows from rep-numbering/timeline-mark context.
btnForgetCsv.addEventListener('click', () => {
  if (loadedRows.length && !confirm(`Forget the ${loadedRows.length} loaded row(s)? They won't be included next time you Save (the file on disk itself is not changed).`)) {
    return;
  }
  loadedRows = [];
  csvFileHandle = null;
  csvFileNameEl.textContent = 'No file loaded — Save will create a new file';
  delete csvFileNameEl.dataset.rawName;
  btnForgetCsv.disabled = true;
  updateRepNumberDefault();
  status('Forgot the loaded file. Save will now create a new file.');
});

/* ---------- Player controls ---------- */

function updateFrameDisplay() {
  const f = frameFromTime(video.currentTime);
  const total = video.duration ? frameFromTime(video.duration) : 0;
  frameDisplay.textContent =
    `Frame ${f} / ${total}   (t = ${video.currentTime.toFixed(3)}s / ${(video.duration || 0).toFixed(3)}s)`;
  if (!seekMainDragging) seekMain.value = video.currentTime;
  updateZoomWindow();
}

// The zoom window is a narrower slice of the timeline (a few seconds
// wide, set by "Window (s)") so dragging lands you precisely on one
// frame — on a long video, each pixel of the main timeline can cover many
// frames, too coarse to hit an exact one. It only re-centers on the
// current frame when that frame actually falls outside its current
// range (e.g. after a big jump elsewhere); otherwise it holds still so a
// drag doesn't visually snap back to the middle the moment you release it.
let zoomInitialized = false;

function updateZoomWindow(forceRecenter) {
  if (zoomDragging || !video.duration) return;
  const cur = video.currentTime;
  const outOfRange = !zoomInitialized || cur < parseFloat(seekZoom.min) || cur > parseFloat(seekZoom.max);
  if (forceRecenter || outOfRange) {
    const win = parseFloat(zoomWindowSize.value) || 2;
    const half = win / 2;
    let lo = cur - half;
    let hi = cur + half;
    if (lo < 0) { hi -= lo; lo = 0; }
    if (hi > video.duration) { lo -= (hi - video.duration); hi = video.duration; lo = Math.max(0, lo); }
    seekZoom.min = lo;
    seekZoom.max = hi;
    seekZoom.step = 1 / getFps();
    zoomInitialized = true;
  }
  seekZoom.value = cur;
}

video.addEventListener('timeupdate', updateFrameDisplay);
video.addEventListener('seeked', updateFrameDisplay);
// Both labels stay in sync with the video's actual state regardless of
// which control (this button, the review panel's, or Space) triggered it.
video.addEventListener('play', () => {
  btnPlayPause.textContent = 'Pause';
  btnReviewPlay.textContent = '⏸ Pause';
});
video.addEventListener('pause', () => {
  btnPlayPause.textContent = 'Play';
  btnReviewPlay.textContent = '▶ Play Rep';
});

// Sliders are blurred after use so the global arrow-key handler (which
// pauses the video and steps exactly one frame) always wins over the
// browser's native "nudge a focused range input by its step" behavior.
seekMain.addEventListener('pointerdown', () => { seekMainDragging = true; cancelPreview(); });
seekMain.addEventListener('pointerup', () => { seekMainDragging = false; seekMain.blur(); });
seekMain.addEventListener('input', () => { video.currentTime = parseFloat(seekMain.value); });

seekZoom.addEventListener('pointerdown', () => { zoomDragging = true; cancelPreview(); });
seekZoom.addEventListener('pointerup', () => { zoomDragging = false; seekZoom.blur(); });
seekZoom.addEventListener('input', () => { video.currentTime = parseFloat(seekZoom.value); });

zoomWindowSize.addEventListener('input', () => updateZoomWindow(true));
fpsInput.addEventListener('input', () => {
  seekMain.step = 1 / getFps();
  updateFrameDisplay();
});

function togglePlayPause() {
  // No cancelPreview() here on purpose: "Check" arms an auto-stop at the
  // rep's End frame but deliberately doesn't auto-play (see openReview) —
  // pressing Play is how you actually start watching it, so this must not
  // cancel the very thing Play is meant to trigger.
  if (video.paused) video.play(); else video.pause();
}

function stepFrame(delta) {
  if (!videoLoaded) return;
  cancelPreview();
  video.pause();
  const f = Math.max(0, frameFromTime(video.currentTime) + delta);
  video.currentTime = Math.min(video.duration || 0, timeFromFrame(f));
}

function jumpSeconds(delta) {
  if (!videoLoaded) return;
  cancelPreview();
  video.currentTime = Math.max(0, Math.min(video.duration || 0, video.currentTime + delta));
}

// "Check" (table below) reviews one saved repetition using the player
// controls right here — no separate mini-player, no scrolling to a
// preview elsewhere. It seeks to the rep's Start frame (paused — see
// openReview) and arms previewEndTime so pressing Play stops exactly at
// its End frame; any manual scrub/step/mark cancels that arm so a
// finished or abandoned preview never lingers and unexpectedly cuts off
// normal playback later.
let previewEndTime = null;
let reviewingRow = null;

function cancelPreview() {
  previewEndTime = null;
}

function stopAtPreviewEnd() {
  if (previewEndTime === null) return false;
  if (video.currentTime < previewEndTime) return false;
  video.pause();
  video.currentTime = previewEndTime; // land exactly on the End frame, not wherever playback happened to be
  previewEndTime = null;
  return true;
}

// timeupdate only fires a handful of times per second — not once per
// frame — so relying on it alone means two things look wrong during
// playback: the frame counter visibly jumps instead of counting up one
// by one (stepping looks smooth only because it explicitly seeks to each
// frame instead), and the preview auto-stop can overshoot the End frame
// by several frames before snapping back, worse at higher speeds.
// Chrome/Edge support requestVideoFrameCallback, which fires once per
// actually rendered frame, fixing both; timeupdate stays registered
// below as the fallback for browsers without it.
if (video.requestVideoFrameCallback) {
  const onVideoFrame = () => {
    updateFrameDisplay();
    if (!stopAtPreviewEnd() && !video.paused) video.requestVideoFrameCallback(onVideoFrame);
  };
  video.addEventListener('play', () => video.requestVideoFrameCallback(onVideoFrame));
}
video.addEventListener('timeupdate', stopAtPreviewEnd);

function renderReviewPanel() {
  if (!reviewingRow) return;
  reviewTitle.textContent = `Checking rep ${reviewingRow.repetition_number} — ${reviewingRow.video_id}`;
  reviewStartInput.value = reviewingRow.first_frame;
  reviewEndInput.value = reviewingRow.last_frame;
}

function openReview(row) {
  if (!videoLoaded) { status('Open the video first.', 'error'); return; }
  if (row.video_id !== videoIdInput.value.trim()) {
    status(`"${row.video_id}" isn't the open video — open it first to check this rep.`, 'error');
    return;
  }
  const startFrame = parseInt(row.first_frame, 10);
  const endFrame = parseInt(row.last_frame, 10);
  if (isNaN(startFrame) || isNaN(endFrame)) { status('This row has invalid frame numbers.', 'error'); return; }

  reviewingRow = row;
  video.pause();
  video.currentTime = timeFromFrame(startFrame);
  previewEndTime = timeFromFrame(endFrame);

  renderReviewPanel();
  reviewPanel.hidden = false;
  reviewPanel.scrollIntoView({ behavior: 'smooth', block: 'center' });
}

function closeReview() {
  reviewingRow = null;
  reviewPanel.hidden = true;
  cancelPreview();
}

// The review panel's own Play button — distinct from the general
// Play/Pause above it. That one just plays/pauses wherever the playhead
// happens to be; this one always (re)starts from this rep's Start frame
// and re-arms the auto-stop, so it still works as "replay this rep" even
// after you've scrubbed elsewhere.
function playReviewedRep() {
  if (!reviewingRow) return;
  const startFrame = parseInt(reviewingRow.first_frame, 10);
  const endFrame = parseInt(reviewingRow.last_frame, 10);
  if (isNaN(startFrame) || isNaN(endFrame)) return;
  video.currentTime = timeFromFrame(startFrame);
  previewEndTime = timeFromFrame(endFrame);
  video.play();
}

// A genuine play/pause toggle, not "always restart from the top": while
// playing, it just pauses wherever the playhead is (previewEndTime is
// left armed, so a plain resume still auto-stops at End). Only restarts
// from the Start frame when there's nothing armed — i.e. it already
// played through to the End and stopped there, where a plain video.play()
// would just run on into whatever comes after this rep.
function toggleReviewPlay() {
  if (!video.paused) { video.pause(); return; }
  if (previewEndTime === null) playReviewedRep();
  else video.play();
}

// Shared by the Start/End number inputs and their +/- nudge buttons:
// updates the reviewed row (same object rendered in the table, so the
// table reflects it immediately), re-arms the preview end if End changed,
// and seeks the player there so the new boundary is visible right away.
function applyReviewFrame(field, value) {
  if (!reviewingRow || isNaN(value)) return;
  reviewingRow[field] = Math.max(0, value);
  if (field === 'last_frame') previewEndTime = timeFromFrame(reviewingRow.last_frame);
  video.pause();
  video.currentTime = timeFromFrame(reviewingRow[field]);
  renderReviewPanel();
  renderTable();
}

function findNextRepForVideo(row) {
  const sameVideo = sessionRows
    .filter(r => r.video_id === row.video_id)
    .sort((a, b) => (parseInt(a.first_frame, 10) || 0) - (parseInt(b.first_frame, 10) || 0));
  const idx = sameVideo.indexOf(row);
  return (idx >= 0 && idx < sameVideo.length - 1) ? sameVideo[idx + 1] : null;
}

btnReviewClose.addEventListener('click', closeReview);
btnReviewPlay.addEventListener('click', toggleReviewPlay);
reviewStartInput.addEventListener('change', () => applyReviewFrame('first_frame', parseInt(reviewStartInput.value, 10)));
reviewEndInput.addEventListener('change', () => applyReviewFrame('last_frame', parseInt(reviewEndInput.value, 10)));
btnReviewStartMinus.addEventListener('click', () => applyReviewFrame('first_frame', (parseInt(reviewingRow.first_frame, 10) || 0) - 1));
btnReviewStartPlus.addEventListener('click', () => applyReviewFrame('first_frame', (parseInt(reviewingRow.first_frame, 10) || 0) + 1));
btnReviewEndMinus.addEventListener('click', () => applyReviewFrame('last_frame', (parseInt(reviewingRow.last_frame, 10) || 0) - 1));
btnReviewEndPlus.addEventListener('click', () => applyReviewFrame('last_frame', (parseInt(reviewingRow.last_frame, 10) || 0) + 1));
btnReviewNext.addEventListener('click', () => {
  const next = findNextRepForVideo(reviewingRow);
  if (!next) { status('No more repetitions for this video.'); return; }
  openReview(next);
});

btnPlayPause.addEventListener('click', togglePlayPause);
btnStepBack.addEventListener('click', () => stepFrame(-1));
btnStepFwd.addEventListener('click', () => stepFrame(1));
btnJumpBack.addEventListener('click', () => jumpSeconds(-1));
btnJumpFwd.addEventListener('click', () => jumpSeconds(1));
speedSelect.addEventListener('change', () => { video.playbackRate = parseFloat(speedSelect.value); });

/* ---------- Marking ---------- */
// One button does double duty: first press marks the start frame and
// flips to "Mark End"; second press marks the end frame and flips back,
// ready for the next rep. "Clear" backs out of an in-progress or
// not-yet-added mark.

function updateMarkDisplays() {
  startDisplay.textContent = 'Start: ' + (pendingStart === null ? '—' : pendingStart);
  endDisplay.textContent = 'End: ' + (pendingEnd === null ? '—' : pendingEnd);
  btnMarkToggle.textContent = markPhase === 'start' ? 'Mark Start (S)' : 'Mark End (S)';
  btnMarkToggle.classList.toggle('recording', markPhase === 'end');
  renderTimelineMarks();
}

function resetMarks() {
  pendingStart = null; pendingEnd = null; markPhase = 'start';
  updateMarkDisplays();
}

// If the video is playing when you mark, freeze it on that exact frame for
// half a second before auto-resuming — a brief visual confirmation of the
// frame you just marked, without permanently interrupting playback. Marking
// while already paused (e.g. after stepping frame-by-frame) is unaffected.
const MARK_PAUSE_MS = 500;
let markPauseTimeout = null;

function toggleMark() {
  if (!videoLoaded) return;
  // Gate here, not just at save time: frame numbers are computed from the
  // current FPS the instant you mark, so marking on a wrong/missing FPS
  // and fixing it afterwards wouldn't retroactively correct them.
  if (!(parseFloat(fpsInput.value) > 0)) { status('Set FPS before marking — frame numbers depend on it.', 'error'); return; }
  cancelPreview();
  const f = frameFromTime(video.currentTime);
  const wasPlaying = !video.paused;
  if (markPauseTimeout) { clearTimeout(markPauseTimeout); markPauseTimeout = null; }
  if (wasPlaying) video.pause();

  if (markPhase === 'start') {
    if (pendingEnd !== null) status('Discarded previous unsaved start/end marks.', 'error');
    pendingStart = f;
    pendingEnd = null;
    markPhase = 'end';
    updateMarkDisplays();
  } else {
    pendingEnd = f;
    markPhase = 'start';
    updateMarkDisplays();
    // Person/exercise/orientation are set up front (section 1), so this
    // is normally ready to save the instant End is marked — no extra
    // click. If something up there was left blank, trySaveRepetition()
    // says so and the change listeners below retry once it's filled in.
    trySaveRepetition();
  }

  if (wasPlaying) {
    markPauseTimeout = setTimeout(() => {
      markPauseTimeout = null;
      video.play();
    }, MARK_PAUSE_MS);
  }
}

// Which of the required, set-up-front fields are still empty right now.
function missingFields() {
  const missing = [];
  if (!videoIdInput.value.trim()) missing.push('Video ID');
  // Not auto-detected from the video file — must be entered explicitly
  // every time, since a silently-wrong default would throw off every
  // frame number computed from it.
  if (!(parseFloat(fpsInput.value) > 0)) missing.push('FPS');
  if (!personInput.value.trim()) missing.push('Person');
  if (!exerciseIdInput.value.trim()) missing.push('Exercise ID');
  if (!currentOrientation()) missing.push('Orientation');
  return missing;
}

// Saves the pending Start/End mark the moment it's complete. If a
// required field was left blank, the marks are kept (nothing is lost)
// and this gets retried automatically as soon as that field is filled in
// (see the 'change' listeners below) — no separate "Add" button needed.
function trySaveRepetition() {
  if (!videoLoaded || pendingStart === null || pendingEnd === null) return;
  if (pendingEnd < pendingStart) { status('End frame is before Start frame — fix it with Clear and re-mark.', 'error'); return; }
  const missing = missingFields();
  if (missing.length > 0) {
    status(`Marked frames ${pendingStart}–${pendingEnd}, but can't save until you fill in: ${missing.join(', ')}.`, 'error');
    return;
  }
  addRepetitionRow();
}

btnMarkToggle.addEventListener('click', toggleMark);
btnClearMarks.addEventListener('click', resetMarks);
btnNewPerson.addEventListener('click', () => {
  personInput.value = '';
  personInput.focus();
  status("Enter the new person's name.");
});
// 'change' (fires on blur/Enter, not per keystroke) so a retry never fires
// on a half-typed value.
exerciseIdInput.addEventListener('change', trySaveRepetition);
personInput.addEventListener('change', trySaveRepetition);
orientationOther.addEventListener('change', trySaveRepetition);

orientationSelect.addEventListener('change', () => {
  orientationOther.hidden = orientationSelect.value !== 'other';
  trySaveRepetition();
});

/* ---------- Add repetition ---------- */

// Each data cell is a plain-looking input you can click into to fix a
// value after the fact (e.g. nudge a Start/Last frame) — it only looks
// like an input on hover/focus (see .cell-input in style.css).
const ROW_NUMBER_FIELDS = ['repetition_number', 'first_frame', 'last_frame'];

function makeCellInput(row, field, isNumber) {
  const input = document.createElement('input');
  input.type = isNumber ? 'number' : 'text';
  input.className = 'cell-input';
  if (isNumber) input.min = field === 'repetition_number' ? '1' : '0';
  input.value = row[field];
  input.addEventListener('change', () => {
    row[field] = isNumber ? (parseInt(input.value, 10) || 0) : input.value.trim();
    input.value = row[field];
    // Keep the review panel in sync if this is the row currently under
    // "Check" — editing it in the table directly should look the same as
    // editing it via the panel's own fields.
    if (row === reviewingRow) {
      if (field === 'last_frame') previewEndTime = timeFromFrame(row.last_frame);
      renderReviewPanel();
    }
    renderTimelineMarks();
  });
  return input;
}

function renderTable() {
  repTableBody.innerHTML = '';
  const overlaps = computeOverlappingRows(sessionRows);

  sessionRows.forEach((r, idx) => {
    const tr = document.createElement('tr');
    if (overlaps.has(r)) {
      tr.classList.add('row-overlap');
      tr.title = 'Overlaps with another repetition for this video — check the First/Last frames.';
    }

    ['video_id', 'repetition_number', 'exercise_id', 'person', 'first_frame', 'last_frame'].forEach(field => {
      const td = document.createElement('td');
      td.appendChild(makeCellInput(r, field, ROW_NUMBER_FIELDS.includes(field)));
      tr.appendChild(td);
    });

    // Computed from First/Last for reference only — never written to the
    // CSV (buildCsvText only ever reads the 10 real schema fields).
    const tdLen = document.createElement('td');
    const first = parseInt(r.first_frame, 10), last = parseInt(r.last_frame, 10);
    tdLen.textContent = (!isNaN(first) && !isNaN(last)) ? String(last - first + 1) : '—';
    tr.appendChild(tdLen);

    const tdOrient = document.createElement('td');
    tdOrient.appendChild(makeCellInput(r, 'front_cam_orientation', false));
    tr.appendChild(tdOrient);

    const tdActions = document.createElement('td');
    tdActions.className = 'row-actions';

    const btnCheck = document.createElement('button');
    btnCheck.textContent = 'Check';
    btnCheck.title = 'Review this repetition here: play/pause, adjust Start/End, step to the next one';
    btnCheck.addEventListener('click', () => openReview(r));
    tdActions.appendChild(btnCheck);

    const btnDel = document.createElement('button');
    btnDel.textContent = 'Delete';
    btnDel.addEventListener('click', () => {
      sessionRows.splice(idx, 1);
      if (reviewingRow === r) closeReview();
      renderTable();
      updateRepNumberDefault();
    });
    tdActions.appendChild(btnDel);

    tr.appendChild(tdActions);

    repTableBody.appendChild(tr);
  });
  btnClearSession.disabled = sessionRows.length === 0;
  overlapWarning.hidden = overlaps.size === 0;
  renderTimelineMarks();
}

function addRepetitionRow() {
  const videoId = videoIdInput.value.trim();
  if (!videoId) { status('Set a Video ID first.', 'error'); return; }
  const repNumber = parseInt(repNumberInput.value, 10) || nextRepNumber(videoId);

  const row = {
    video_id: videoId,
    repetition_number: repNumber,
    exercise_id: exerciseIdInput.value.trim(),
    person: personInput.value.trim(),
    first_frame: pendingStart,
    last_frame: pendingEnd,
    front_cam_orientation: currentOrientation(),
    correctness_score: '',
    error_description: '',
    video_front_problem: ''
  };
  sessionRows.push(row);
  renderTable();

  resetMarks();
  updateRepNumberDefault();
  status(`Added rep ${repNumber} for ${videoId} (frames ${row.first_frame}–${row.last_frame}).`, 'success');
}

// Clears only this session's newly-added rows (the table above) — an
// appended file's already-loaded rows are untouched; use "Forget Loaded
// File" (section 2) to drop those instead.
btnClearSession.addEventListener('click', () => {
  if (!sessionRows.length) return;
  if (!confirm(`Remove all ${sessionRows.length} repetition(s) added this session? This can't be undone.`)) return;
  sessionRows = [];
  closeReview();
  renderTable();
  resetMarks();
  updateRepNumberDefault();
  status('Cleared this session\'s repetitions.');
});

/* ---------- Save ---------- */

function suggestedCsvName() {
  const existing = csvFileNameEl.dataset.rawName;
  if (existing) return existing;
  const id = videoIdInput.value.trim();
  return (id || 'segmentation') + '.csv';
}

function downloadTextAsFile(text, filename) {
  const blob = new Blob([text], { type: 'text/csv;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

btnSave.addEventListener('click', async () => {
  const allRows = loadedRows.concat(sessionRows);
  if (allRows.length === 0) { status('Nothing to save yet — add at least one repetition.'); return; }
  const text = buildCsvText(allRows);

  if (csvFileHandle) {
    try {
      if ((await csvFileHandle.queryPermission({ mode: 'readwrite' })) !== 'granted') {
        const perm = await csvFileHandle.requestPermission({ mode: 'readwrite' });
        if (perm !== 'granted') throw new Error('write permission denied');
      }
      const writable = await csvFileHandle.createWritable();
      await writable.write(text);
      await writable.close();
      status(`Saved ${allRows.length} rows to ${csvFileHandle.name}.`, 'success');
      return;
    } catch (err) {
      status('Could not write to file (' + err.message + ') — falling back to download.', 'error');
    }
  }

  if (window.showSaveFilePicker) {
    try {
      const handle = await window.showSaveFilePicker({
        suggestedName: suggestedCsvName(),
        types: [{ description: 'CSV files', accept: { 'text/csv': ['.csv'] } }]
      });
      csvFileHandle = handle;
      const writable = await handle.createWritable();
      await writable.write(text);
      await writable.close();
      csvFileNameEl.textContent = `${handle.name} (${allRows.length} rows)`;
      csvFileNameEl.dataset.rawName = handle.name;
      status(`Saved ${allRows.length} rows to ${handle.name}.`, 'success');
      return;
    } catch (err) {
      if (err.name === 'AbortError') { status('Save cancelled.'); return; }
      status('Save failed (' + err.message + ') — falling back to download.', 'error');
    }
  }

  const filename = suggestedCsvName();
  downloadTextAsFile(text, filename);
  status(`Downloaded ${allRows.length} rows as ${filename}. Your browser can't save in place, so replace the old file manually if you were appending.`, 'success');
});

/* ---------- Keyboard shortcuts ---------- */

document.addEventListener('keydown', (e) => {
  const tag = (document.activeElement && document.activeElement.tagName) || '';
  if (['INPUT', 'SELECT', 'TEXTAREA'].includes(tag)) return;
  if (!videoLoaded) return;

  switch (e.key) {
    case ' ':
      e.preventDefault(); togglePlayPause(); break;
    case 'ArrowLeft':
      e.preventDefault(); e.shiftKey ? jumpSeconds(-1) : stepFrame(-1); break;
    case 'ArrowRight':
      e.preventDefault(); e.shiftKey ? jumpSeconds(1) : stepFrame(1); break;
    case 's': case 'S':
      toggleMark(); break;
    default:
      break;
  }
});

/* ---------- Init ---------- */

status(window.showSaveFilePicker
  ? 'Ready. Open a video to begin.'
  : 'Ready. Note: your browser lacks the File System Access API, so Save will always download a new file instead of writing in place.');
