<?php

declare(strict_types=1);

require_once __DIR__ . '/lib/helper.php';

load_env(dirname(__DIR__) . '/.env');

// ── Download handler ───────────────────────────────────────────────────────
if (isset($_GET['download'], $_GET['job'])) {
    $job_id   = preg_replace('/[^a-f0-9]/', '', $_GET['job']);
    $job_file = __DIR__ . '/jobs/' . $job_id . '.json';

    if ($job_id && file_exists($job_file)) {
        $job = json_decode(file_get_contents($job_file), true);
        $out = $job['output_file'] ?? '';
        if ($out && file_exists($out)) {
            header('Content-Type: text/csv; charset=UTF-8');
            header('Content-Disposition: attachment; filename="seo-meta-translated.csv"');
            header('Content-Length: ' . filesize($out));
            readfile($out);
            exit;
        }
    }
    header('Location: index.php');
    exit;
}

// ── Validate job param ─────────────────────────────────────────────────────
$job_id = preg_replace('/[^a-f0-9]/', '', $_GET['job'] ?? '');
if (!$job_id) {
    header('Location: index.php');
    exit;
}

$lang_label = get_language_name('EN'); // placeholder, updated client-side

$status_labels = [
    'deepl_ok'        => ['label' => 'DeepL',     'class' => 'badge-blue'],
    'openai_ok'       => ['label' => 'OpenAI',    'class' => 'badge-green'],
    'gemini_ok'       => ['label' => 'Gemini',    'class' => 'badge-purple'],
    'groq_ok'         => ['label' => 'Groq',      'class' => 'badge-orange'],
    'meaning_fallback'=> ['label' => 'Length OK', 'class' => 'badge-orange'],
    'fallback_trim'   => ['label' => 'Trimmed',   'class' => 'badge-red'],
    'best_available'  => ['label' => 'Best fit',  'class' => 'badge-orange'],
    'failed'          => ['label' => 'Failed',    'class' => 'badge-red'],
    'skipped'         => ['label' => 'Skipped',   'class' => 'badge-gray'],
];

?>
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Processing — SEO Meta Tool</title>
<style>
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; background: #f5f7fa; color: #1a1a2e; }
  .header { background: #1a1a2e; color: #fff; padding: 1rem 2rem; display: flex; align-items: center; justify-content: space-between; }
  .header h1 { font-size: 1.2rem; font-weight: 600; }
  .container { max-width: 1200px; margin: 2rem auto; padding: 0 1.5rem; }
  .card { background: #fff; border-radius: 10px; box-shadow: 0 2px 8px rgba(0,0,0,.08); padding: 1.5rem; margin-bottom: 1.5rem; }

  /* Loading state */
  .loading-wrap { text-align: center; padding: 3rem 1rem; }
  .spinner { width: 52px; height: 52px; border: 5px solid #e9ecef; border-top-color: #3b5bdb; border-radius: 50%; animation: spin .8s linear infinite; margin: 0 auto 1.5rem; }
  @keyframes spin { to { transform: rotate(360deg); } }
  .loading-title { font-size: 1.2rem; font-weight: 600; margin-bottom: .5rem; }
  .loading-sub   { font-size: .9rem; color: #666; }
  .progress-bar  { height: 4px; background: #e9ecef; border-radius: 2px; margin-top: 1.5rem; overflow: hidden; }
  .progress-fill { height: 100%; background: #3b5bdb; border-radius: 2px; animation: indeterminate 1.5s ease-in-out infinite; width: 30%; }
  @keyframes indeterminate { 0%{transform:translateX(-100%)} 100%{transform:translateX(400%)} }

  /* Results */
  .summary { display: flex; gap: 1.5rem; flex-wrap: wrap; }
  .stat { flex: 1; min-width: 140px; text-align: center; padding: 1rem; background: #f0f4ff; border-radius: 8px; }
  .stat-num { font-size: 2rem; font-weight: 700; color: #3b5bdb; }
  .stat-label { font-size: .8rem; color: #555; margin-top: .2rem; }
  .btn { display: inline-flex; align-items: center; gap: .4rem; padding: .6rem 1.2rem; border-radius: 6px; font-size: .9rem; font-weight: 500; text-decoration: none; cursor: pointer; border: none; }
  .btn-primary { background: #3b5bdb; color: #fff; }
  .btn-primary:hover { background: #2f4ac7; }
  .btn-secondary { background: #e9ecef; color: #333; }
  .btn-secondary:hover { background: #dee2e6; }
  .actions { display: flex; gap: .8rem; margin-bottom: 1.5rem; }
  table { width: 100%; border-collapse: collapse; font-size: .85rem; }
  th { background: #f8f9fa; text-align: left; padding: .7rem .9rem; font-weight: 600; color: #444; border-bottom: 2px solid #e9ecef; white-space: nowrap; }
  td { padding: .65rem .9rem; border-bottom: 1px solid #f0f0f0; vertical-align: top; }
  tr:hover td { background: #fafbff; }
  .badge { display: inline-block; padding: .15rem .55rem; border-radius: 20px; font-size: .72rem; font-weight: 600; }
  .badge-blue   { background: #dbe4ff; color: #3b5bdb; }
  .badge-purple { background: #f3d9fa; color: #862e9c; }
  .badge-orange { background: #ffe8cc; color: #d9480f; }
  .badge-red    { background: #ffe3e3; color: #c92a2a; }
  .badge-gray   { background: #f1f3f5; color: #666; }
  .badge-green  { background: #d3f9d8; color: #2b8a3e; }
  .len { font-size: .75rem; color: #888; margin-top: .2rem; }
  .len.valid   { color: #2b8a3e; }
  .len.invalid { color: #c92a2a; }
  .text-cell { max-width: 260px; }
  .text-cell span { display: block; line-height: 1.4; }
  .error-box { background: #fff5f5; border: 1px solid #ffc9c9; border-radius: 8px; padding: 1rem 1.2rem; color: #c92a2a; }
  .section-title { font-size: 1rem; font-weight: 600; margin-bottom: 1rem; color: #333; }
  .overflow-x { overflow-x: auto; }
  #results-section { display: none; }
  #error-section   { display: none; }
</style>
</head>
<body>

<div class="header">
  <h1>SEO Meta Translation Tool</h1>
  <span id="lang-label" style="font-size:.85rem;opacity:.7"></span>
</div>

<div class="container">

  <!-- Loading state -->
  <div id="loading-section" class="card">
    <div class="loading-wrap">
      <div class="spinner"></div>
      <div class="loading-title">Processing your file&hellip;</div>
      <div class="loading-sub">Translating and rewriting meta tags. This may take a minute.</div>
      <div class="progress-bar"><div class="progress-fill"></div></div>
    </div>
  </div>

  <!-- Error state -->
  <div id="error-section" class="card">
    <div class="error-box" id="error-message"></div>
    <div style="margin-top:1rem">
      <a href="index.php" class="btn btn-secondary">← Back</a>
    </div>
  </div>

  <!-- Results state -->
  <div id="results-section">
    <div class="card">
      <div class="summary">
        <div class="stat"><div class="stat-num" id="stat-total">0</div><div class="stat-label">Rows processed</div></div>
        <div class="stat"><div class="stat-num" id="stat-titles">0</div><div class="stat-label">Valid titles</div></div>
        <div class="stat"><div class="stat-num" id="stat-descs">0</div><div class="stat-label">Valid descriptions</div></div>
        <div class="stat"><div class="stat-num" id="stat-trimmed">0</div><div class="stat-label">Hard-trimmed rows</div></div>
      </div>
    </div>

    <div class="actions">
      <a id="download-btn" href="#" class="btn btn-primary">⬇ Download output.csv</a>
      <a href="index.php" class="btn btn-secondary">← Process another file</a>
    </div>

    <div class="card">
      <div class="section-title">Results</div>
      <div class="overflow-x">
        <table>
          <thead>
            <tr>
              <th>#</th><th>Lang</th><th>Original Title</th><th>New Title</th>
              <th>Alt Title</th><th>Title Status</th><th>Original Description</th>
              <th>New Description</th><th>Alt Description</th><th>Desc Status</th>
            </tr>
          </thead>
          <tbody id="results-tbody"></tbody>
        </table>
      </div>
    </div>
  </div>

</div>

<script>
const JOB_ID = <?= json_encode($job_id) ?>;
const POLL_MS = 2000;
const TIMEOUT_MS = 600000; // 10 min

const STATUS_LABELS = <?= json_encode($status_labels) ?>;

const $ = id => document.getElementById(id);

function badgeHtml(status) {
  const s = STATUS_LABELS[status] || { label: status, class: 'badge-gray' };
  return `<span class="badge ${s.class}">${s.label}</span>`;
}

function lenHtml(len, valid) {
  return `<div class="len ${valid ? 'valid' : 'invalid'}">${len} chars ${valid ? '✓' : '✗'}</div>`;
}

function renderResults(job) {
  const r = job.result;
  $('loading-section').style.display = 'none';

  $('lang-label').textContent = 'Job complete';
  $('download-btn').href = 'process.php?download=1&job=' + JOB_ID;

  const rows  = r.rows || [];
  let validT  = 0, validD = 0, trimmed = 0;

  rows.forEach(row => {
    if (row.result_title.valid)  validT++;
    if (row.result_desc.valid)   validD++;
    if (row.result_title.status === 'fallback_trim' || row.result_desc.status === 'fallback_trim') trimmed++;
  });

  $('stat-total').textContent   = r.processed;
  $('stat-titles').textContent  = validT;
  $('stat-descs').textContent   = validD;
  $('stat-trimmed').textContent = trimmed;

  const tbody = $('results-tbody');
  tbody.innerHTML = rows.map((row, i) => {
    const rt = row.result_title;
    const rd = row.result_desc;

    const altT    = row.alt_title || '';
    const altTLen = [...altT].length;
    const altTOk  = altTLen >= 40 && altTLen <= 46;
    const altTCell = altT
      ? `<span>${esc(altT)}</span>${lenHtml(altTLen, altTOk)}`
      : `<span style="color:#bbb;font-size:.8rem;">—</span>`;

    const altD    = row.alt_desc || '';
    const altDLen = [...altD].length;
    const altDOk  = altDLen >= 150 && altDLen <= 155;
    const altDCell = altD
      ? `<span>${esc(altD)}</span>${lenHtml(altDLen, altDOk)}`
      : `<span style="color:#bbb;font-size:.8rem;">—</span>`;

    return `<tr>
      <td>${i + 1}</td>
      <td><span class="badge badge-blue" style="font-size:.7rem">${esc(row.language || '')}</span></td>
      <td class="text-cell"><span>${esc(row.original_title)}</span></td>
      <td class="text-cell"><span>${esc(rt.text)}</span>${rt.text ? lenHtml(rt.len, rt.valid) : ''}</td>
      <td class="text-cell">${altTCell}</td>
      <td>${badgeHtml(rt.status)}<br><small style="color:#aaa">${rt.attempts}× API</small></td>
      <td class="text-cell"><span>${esc(row.original_desc)}</span></td>
      <td class="text-cell"><span>${esc(rd.text)}</span>${rd.text ? lenHtml(rd.len, rd.valid) : ''}</td>
      <td class="text-cell">${altDCell}</td>
      <td>${badgeHtml(rd.status)}<br><small style="color:#aaa">${rd.attempts}× API</small></td>
    </tr>`;
  }).join('');

  $('results-section').style.display = 'block';
}

function showError(msg) {
  $('loading-section').style.display = 'none';
  $('error-message').textContent = msg;
  $('error-section').style.display = 'block';
}

function esc(str) {
  return String(str)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}

let elapsed = 0;
const timer = setInterval(async () => {
  elapsed += POLL_MS;
  if (elapsed >= TIMEOUT_MS) {
    clearInterval(timer);
    showError('Processing timed out. Please try again.');
    return;
  }

  try {
    const res  = await fetch('status.php?job=' + JOB_ID);
    const job  = await res.json();

    if (job.status === 'done') {
      clearInterval(timer);
      renderResults(job);
    } else if (job.status === 'error') {
      clearInterval(timer);
      showError(job.error || 'An error occurred during processing.');
    } else if (job.status === 'not_found') {
      clearInterval(timer);
      showError('Job not found. Please upload your file again.');
    }
    // 'pending' or 'processing' → keep polling
  } catch (e) {
    // network glitch — keep polling
  }
}, POLL_MS);
</script>
</body>
</html>
