<?php session_start(); ?>
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>SEO Meta Translation Tool</title>
<style>
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
    background: #f5f7fa;
    color: #1a1a2e;
    min-height: 100vh;
    display: flex;
    flex-direction: column;
  }
  .header {
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    color: #fff;
    text-align: center;
    padding: 1.4rem 2rem 1.2rem;
  }
  .header h1 { font-size: 2rem; font-weight: 700; margin-bottom: .4rem; }
  .header p  { font-size: 1rem; opacity: .9; }
  .nav-tabs  { display: flex; justify-content: center; gap: 8px; margin-top: 14px; }
  .nav-tab {
    padding: 8px 20px;
    border-radius: 20px;
    text-decoration: none;
    font-size: 14px;
    font-weight: 600;
    color: rgba(255,255,255,.8);
    background: rgba(255,255,255,.15);
    border: 2px solid transparent;
    transition: background .2s;
  }
  .nav-tab:hover { background: rgba(255,255,255,.25); color: #fff; }
  .nav-tab.active { background: white; color: #667eea; border-color: white; }

  .main {
    flex: 1;
    display: flex;
    align-items: center;
    justify-content: center;
    padding: 2rem 1rem;
  }

  .card {
    background: #fff;
    border-radius: 12px;
    box-shadow: 0 4px 20px rgba(0,0,0,.1);
    padding: 2rem 2.5rem;
    width: 100%;
    max-width: 520px;
  }

  .card-title {
    font-size: 1.3rem;
    font-weight: 700;
    margin-bottom: .4rem;
  }
  .card-sub {
    font-size: .85rem;
    color: #666;
    margin-bottom: 1.8rem;
    line-height: 1.5;
  }

  .field { margin-bottom: 1.2rem; }
  label  { display: block; font-size: .85rem; font-weight: 600; color: #333; margin-bottom: .4rem; }

  select, input[type="file"] {
    width: 100%;
    padding: .6rem .9rem;
    border: 1.5px solid #d0d7e2;
    border-radius: 7px;
    font-size: .9rem;
    color: #1a1a2e;
    background: #fff;
    outline: none;
    transition: border-color .15s;
  }
  select:focus, input[type="file"]:focus {
    border-color: #3b5bdb;
  }

  .drop-zone {
    border: 2px dashed #c5cfe8;
    border-radius: 8px;
    padding: 1.6rem 1rem;
    text-align: center;
    cursor: pointer;
    transition: border-color .2s, background .2s;
    position: relative;
  }
  .drop-zone:hover, .drop-zone.dragover {
    border-color: #3b5bdb;
    background: #f0f4ff;
  }
  .drop-zone input[type="file"] {
    position: absolute; inset: 0; opacity: 0; cursor: pointer; width: 100%; height: 100%;
    border: none; padding: 0;
  }
  .drop-icon { font-size: 2rem; margin-bottom: .4rem; }
  .drop-text { font-size: .85rem; color: #555; }
  .drop-hint { font-size: .75rem; color: #aaa; margin-top: .3rem; }
  .file-name  { font-size: .85rem; color: #3b5bdb; margin-top: .5rem; font-weight: 500; display: none; }
  .btn-template {
    display: inline-flex; align-items: center; gap: .3rem;
    padding: .35rem .85rem; border-radius: 6px; font-size: .78rem; font-weight: 600;
    text-decoration: none; color: #3b5bdb;
    background: #eef1ff; border: 1.5px solid #c5cfe8;
    white-space: nowrap; transition: background .15s;
  }
  .btn-template:hover { background: #dbe4ff; }

  .btn-submit {
    width: 100%;
    padding: .8rem;
    background: #3b5bdb;
    color: #fff;
    border: none;
    border-radius: 8px;
    font-size: 1rem;
    font-weight: 600;
    cursor: pointer;
    transition: background .15s;
    margin-top: .5rem;
  }
  .btn-submit:hover { background: #2f4ac7; }
  .btn-submit:disabled { background: #a5b4fc; cursor: not-allowed; }

  .error-box {
    background: #fff5f5;
    border: 1px solid #ffc9c9;
    color: #c92a2a;
    border-radius: 7px;
    padding: .75rem 1rem;
    font-size: .85rem;
    margin-bottom: 1.2rem;
  }

  .info-box {
    background: #f0f4ff;
    border-radius: 8px;
    padding: 1rem 1.2rem;
    font-size: .8rem;
    color: #3b5bdb;
    margin-top: 1.5rem;
    line-height: 1.6;
  }
  .info-box strong { display: block; margin-bottom: .3rem; color: #1a1a2e; }

  .constraints { display: flex; gap: 1rem; margin-top: .5rem; }
  .constraint {
    flex: 1;
    background: #fff;
    border-radius: 6px;
    padding: .5rem .8rem;
    font-size: .78rem;
    color: #444;
    border: 1px solid #d0d7e2;
  }
  .constraint span { display: block; font-weight: 700; color: #3b5bdb; font-size: .95rem; }
</style>
</head>
<body>

<div class="header">
  <h1>SEO Meta Translation Tool</h1>
  <p>Translate &amp; enforce SEO character limits via DeepL + AI rewriting</p>
  <div class="nav-tabs">
    <a href="../index.php" class="nav-tab">DOCX Generator</a>
    <a href="../meta-extractor.php" class="nav-tab">Meta Extractor</a>
    <a href="index.php" class="nav-tab active">Meta Translator</a>
    <a href="../lang-generator.php" class="nav-tab">Language Tab Generator</a>
    <a href="../cyrillic-cleaner.php" class="nav-tab">Cyrillic Cleaner</a>
  </div>
</div>

<div class="main">
  <div class="card">
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:.4rem;">
      <div class="card-title" style="margin-bottom:0">Upload CSV</div>
      <a href="download-template.php" class="btn-template" title="Download a sample CSV to see the expected format">
        ⬇ Download Template
      </a>
    </div>
    <p class="card-sub">
      Upload a CSV with <code>meta_title</code> and/or <code>meta_description</code> columns.
      Each field is translated, validated, and auto-rewritten if needed.
    </p>

    <?php
    if (!empty($_SESSION['error'])):
    ?>
    <div class="error-box"><?= htmlspecialchars($_SESSION['error']) ?></div>
    <?php
      unset($_SESSION['error']);
    endif;
    ?>

    <form action="upload.php" method="POST" enctype="multipart/form-data" id="uploadForm">

      <div class="field">
        <label for="target_lang">Target Language</label>
        <select name="target_lang" id="target_lang" required>
          <option value="">— Select language —</option>
          <option value="FR">French (FR)</option>
          <option value="DE">German (DE)</option>
          <option value="ES">Spanish (ES)</option>
          <option value="IT">Italian (IT)</option>
          <option value="NL">Dutch (NL)</option>
          <option value="PT">Portuguese (PT)</option>
          <option value="PT-BR">Portuguese — Brazil (PT-BR)</option>
          <option value="PL">Polish (PL)</option>
          <option value="RU">Russian (RU)</option>
          <option value="SV">Swedish (SV)</option>
          <option value="DA">Danish (DA)</option>
          <option value="FI">Finnish (FI)</option>
          <option value="TR">Turkish (TR)</option>
          <option value="CS">Czech (CS)</option>
          <option value="RO">Romanian (RO)</option>
          <option value="JA">Japanese (JA)</option>
          <option value="ZH">Chinese Simplified (ZH)</option>
        </select>
      </div>

      <div class="field">
        <label>CSV File</label>
        <div class="drop-zone" id="dropZone">
          <input type="file" name="csv_file" id="csvFile" accept=".csv" required>
          <div class="drop-icon">📄</div>
          <div class="drop-text">Click to browse or drag &amp; drop</div>
          <div class="drop-hint">Accepts .csv files up to 10 MB</div>
          <div class="file-name" id="fileName"></div>
        </div>
      </div>

      <button type="submit" class="btn-submit" id="submitBtn">Translate &amp; Process</button>
    </form>

    <div class="info-box">
      <strong>SEO Constraints enforced:</strong>
      <div class="constraints">
        <div class="constraint">Meta Title<span>40–46 chars</span></div>
        <div class="constraint">Meta Description<span>150–155 chars</span></div>
      </div>
    </div>
  </div>
</div>

<script>
const dropZone  = document.getElementById('dropZone');
const csvFile   = document.getElementById('csvFile');
const fileName  = document.getElementById('fileName');
const submitBtn = document.getElementById('submitBtn');

csvFile.addEventListener('change', () => {
  if (csvFile.files.length) {
    fileName.textContent = '✓ ' + csvFile.files[0].name;
    fileName.style.display = 'block';
  }
});

['dragover','dragleave','drop'].forEach(evt => {
  dropZone.addEventListener(evt, e => {
    e.preventDefault();
    dropZone.classList.toggle('dragover', evt === 'dragover');
  });
});

document.getElementById('uploadForm').addEventListener('submit', () => {
  submitBtn.disabled = true;
  submitBtn.textContent = 'Processing…';
});
</script>
</body>
</html>
