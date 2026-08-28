<?php
/**
 * Cyrillic Word Cleaner
 * Route: /cyrillic-cleaner.php
 *
 * Cleans corrupted English text containing Cyrillic / lookalike Unicode characters.
 *
 * Two layers of detection:
 *   Layer 1 - word-level dictionary built from crilic-wordss.csv (see cyrillic-dictionary.php)
 *   Layer 2 - generic character-level lookalike mapping, for corrupted words that
 *             are not present in the CSV at all.
 *
 * All replacement happens locally in the browser. No API calls, no rewriting.
 */
?>
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Cyrillic Word Cleaner</title>
<style>
    * { margin: 0; padding: 0; box-sizing: border-box; }

    body {
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        min-height: 100vh;
        padding: 20px;
        color: #333;
    }

    .header { text-align: center; color: #fff; margin-bottom: 20px; }
    .header h1 { font-size: 32px; margin-bottom: 8px; }
    .header p { font-size: 16px; opacity: .9; }
    .nav-tabs { display: flex; justify-content: center; gap: 8px; margin-top: 14px; flex-wrap: wrap; }
    .nav-tab {
        padding: 8px 20px; border-radius: 20px; text-decoration: none; font-size: 14px;
        font-weight: 600; color: rgba(255,255,255,.85); background: rgba(255,255,255,.15);
        border: 2px solid transparent; transition: all .2s;
    }
    .nav-tab:hover { background: rgba(255,255,255,.28); color: #fff; }
    .nav-tab.active { background: #fff; color: #667eea; border-color: #fff; }

    .main { max-width: 1400px; margin: 0 auto; }

    .card {
        background: #fff; border-radius: 12px; padding: 22px;
        box-shadow: 0 10px 30px rgba(0,0,0,.15); margin-bottom: 20px;
    }
    .card-title { font-size: 18px; font-weight: 700; color: #2d3748; margin-bottom: 14px; }

    /* ---- Toolbar ---- */
    .toolbar { display: flex; flex-wrap: wrap; gap: 10px; align-items: center; }
    .toolbar .spacer { flex: 1 1 auto; }

    /* Each pane owns its own actions, so the two columns end at the same height. */
    .pane-actions { display: flex; justify-content: flex-end; gap: 10px; margin-top: 12px; }

    .btn {
        padding: 11px 22px; border: none; border-radius: 8px; font-size: 14px; font-weight: 600;
        cursor: pointer; transition: all .2s; white-space: nowrap;
    }
    .btn:disabled { opacity: .5; cursor: not-allowed; }
    .btn-primary { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: #fff; }
    .btn-primary:not(:disabled):hover { transform: translateY(-1px); box-shadow: 0 6px 16px rgba(102,126,234,.4); }
    .btn-secondary { background: #fff; color: #667eea; border: 2px solid #667eea; }
    .btn-secondary:not(:disabled):hover { background: #667eea; color: #fff; }
    .btn-ghost { background: #f1f3f9; color: #4a5568; border: 2px solid #e2e8f0; }
    .btn-ghost:not(:disabled):hover { background: #e2e8f0; }
    .btn-copied { background: #28a745 !important; color: #fff !important; border-color: #28a745 !important; }

    .options { display: flex; flex-wrap: wrap; gap: 8px 18px; margin-top: 16px;
               padding-top: 16px; border-top: 1px solid #edf0f7; }
    .opt { display: flex; align-items: center; gap: 7px; font-size: 13.5px; color: #4a5568;
           cursor: pointer; user-select: none; }
    .opt input { width: 16px; height: 16px; accent-color: #667eea; cursor: pointer; flex: none; }
    .opt .hint { color: #a0aec0; font-size: 12px; }

    /* ---- Panes ---- */
    .panes { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }
    .pane { display: flex; flex-direction: column; min-width: 0; }
    .pane-head {
        display: flex; justify-content: space-between; align-items: baseline;
        margin-bottom: 8px; gap: 10px;
    }
    .pane-head label { font-size: 14px; font-weight: 700; color: #2d3748; }
    .pane-head .meta { font-size: 12px; color: #a0aec0; white-space: nowrap; }
    textarea {
        width: 100%; min-height: 340px; padding: 14px; border: 2px solid #e2e8f0; border-radius: 10px;
        font-family: ui-monospace, SFMono-Regular, 'SF Mono', Menlo, Consolas, 'Liberation Mono', monospace;
        font-size: 14px; line-height: 1.65; resize: vertical; background: #fff; color: #2d3748;
        flex: 1 1 auto;   /* fill the pane so both columns line up */
    }
    textarea:focus { outline: none; border-color: #667eea; }
    #output { background: #f8fafc; }

    /* ---- Rich-text editor toolbar ---- */
    .editor-toolbar {
        display: flex; flex-wrap: wrap; align-items: center; gap: 4px;
        padding: 7px 9px; border: 2px solid #e2e8f0; border-bottom: none;
        border-radius: 10px 10px 0 0; background: #f8fafc;
    }
    .editor-toolbar select {
        padding: 5px 8px; border: 1px solid #e2e8f0; border-radius: 6px; background: #fff;
        font-size: 13px; color: #2d3748; cursor: pointer; font-family: inherit;
    }
    .editor-toolbar select:focus { outline: none; border-color: #667eea; }
    .tb-btn {
        min-width: 30px; padding: 5px 9px; border: 1px solid transparent; border-radius: 6px;
        background: transparent; color: #4a5568; font-size: 13px; cursor: pointer;
        font-family: inherit; line-height: 1.4;
    }
    .tb-btn:hover { background: #e9edf5; }
    .tb-btn.on { background: #667eea; color: #fff; border-color: #667eea; }
    .tb-sep { width: 1px; height: 20px; background: #e2e8f0; margin: 0 3px; }

    /* The toolbar sits directly on top of the editor, so square off the seam. */
    .editor-toolbar + .editor { border-radius: 0 0 10px 10px; }

    /* ---- Rich-text editor panes ---- */
    .editor {
        width: 100%; min-height: 340px; max-height: 70vh; padding: 14px;
        border: 2px solid #e2e8f0; border-radius: 10px; background: #fff; color: #2d3748;
        font-size: 15px; line-height: 1.7; overflow-y: auto; flex: 1 1 auto;
    }
    .editor:focus { outline: none; border-color: #667eea; }
    .editor.out { background: #f8fafc; }
    .editor:empty:before {
        content: attr(data-placeholder); color: #a0aec0; pointer-events: none;
    }
    /* Make pasted documents look like documents, without overriding inline styles. */
    .editor h1, .editor h2, .editor h3, .editor h4 { margin: 1em 0 .45em; line-height: 1.3; }
    .editor h1 { font-size: 1.7em; } .editor h2 { font-size: 1.4em; }
    .editor h3 { font-size: 1.18em; } .editor h4 { font-size: 1.05em; }
    .editor p { margin: 0 0 .85em; }
    .editor ul, .editor ol { margin: 0 0 .85em; padding-left: 1.6em; }
    .editor li { margin-bottom: .3em; }
    .editor a { color: #667eea; }
    .editor blockquote {
        margin: 0 0 .85em; padding-left: 14px; border-left: 3px solid #e2e8f0; color: #4a5568;
    }
    .editor img { max-width: 100%; height: auto; }
    .editor table { min-width: 0; font-size: .95em; margin-bottom: .85em; }
    .editor > *:first-child { margin-top: 0; }
    .editor > *:last-child  { margin-bottom: 0; }

    /* ---- Mode selector ---- */
    .modes { display: flex; flex-wrap: wrap; align-items: center; gap: 8px 18px;
             margin-top: 16px; padding-top: 16px; border-top: 1px solid #edf0f7; }
    .modes-label { font-size: 13px; font-weight: 700; color: #2d3748;
                   text-transform: uppercase; letter-spacing: .4px; }
    .mode { display: flex; align-items: center; gap: 7px; font-size: 13.5px;
            color: #4a5568; cursor: pointer; user-select: none; }
    .mode input { width: 16px; height: 16px; accent-color: #667eea; cursor: pointer; flex: none; }
    .mode .hint { color: #a0aec0; font-size: 12px; }

    /* ---- Stats ---- */
    .stats { display: grid; grid-template-columns: repeat(4, 1fr); gap: 14px; }
    .stat { background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 10px; padding: 14px 16px; }
    .stat .num { font-size: 26px; font-weight: 800; color: #667eea; line-height: 1.1; }
    .stat .lbl { font-size: 12.5px; color: #718096; margin-top: 4px; }
    .stat.warn .num { color: #d69e2e; }
    .stat.warn { background: #fffaf0; border-color: #f6e05e; }

    .banner { border-radius: 10px; padding: 13px 16px; font-size: 14.5px; font-weight: 600;
              margin-top: 16px; display: flex; gap: 10px; align-items: flex-start; }
    .banner.ok   { background: #f0fff4; color: #22703a; border: 1px solid #9ae6b4; }
    .banner.warn { background: #fffaf0; color: #8a5a00; border: 1px solid #f6e05e; }
    .banner.info { background: #ebf4ff; color: #2c5282; border: 1px solid #bee3f8; }
    .banner.err  { background: #fff5f5; color: #9b2c2c; border: 1px solid #feb2b2; }
    .banner .body { font-weight: 500; }
    .banner .body strong { font-weight: 700; }

    .codepoints { display: flex; flex-wrap: wrap; gap: 7px; margin-top: 9px; }
    .cp {
        background: #fff; border: 1px solid #e2e8f0; border-radius: 6px; padding: 3px 9px;
        font-family: ui-monospace, Menlo, Consolas, monospace; font-size: 12.5px; color: #4a5568;
    }
    .cp b { color: #c53030; font-size: 14px; }
    .cp .n { color: #a0aec0; }

    /* ---- Change log ---- */
    .table-wrap { overflow-x: auto; -webkit-overflow-scrolling: touch; }
    table { width: 100%; border-collapse: collapse; font-size: 14px; min-width: 480px; }
    th, td { padding: 9px 12px; text-align: left; border-bottom: 1px solid #edf0f7; }
    th { background: #f8fafc; color: #4a5568; font-size: 12px; text-transform: uppercase;
         letter-spacing: .4px; position: sticky; top: 0; }
    td.mono { font-family: ui-monospace, Menlo, Consolas, monospace; }
    td.num, th.num { text-align: right; }
    .tag { display: inline-block; padding: 2px 9px; border-radius: 20px; font-size: 11.5px; font-weight: 700; }
    .tag.csv     { background: #e9e5ff; color: #553c9a; }
    .tag.unicode { background: #e6fffa; color: #285e61; }
    .tag.invis   { background: #fff5f5; color: #9b2c2c; }
    .from { color: #c53030; }
    .to   { color: #22703a; }

    .muted { color: #718096; font-size: 13px; }
    .dict-line { font-size: 13px; color: #718096; display: flex; flex-wrap: wrap;
                 gap: 6px 14px; align-items: center; }
    .dict-line b { color: #4a5568; }
    .dict-line a { color: #667eea; font-weight: 600; text-decoration: none; cursor: pointer; }
    .dict-line a:hover { text-decoration: underline; }
    .dict-line .skipped { color: #b7791f; font-weight: 600; cursor: help; }
    .dot { width: 8px; height: 8px; border-radius: 50%; background: #cbd5e0; flex: none; }
    .dot.ok { background: #38a169; }
    .dot.err { background: #e53e3e; }

    .hidden { display: none !important; }

    @media (max-width: 900px) {
        body { padding: 12px; }
        .header h1 { font-size: 24px; }
        .header p { font-size: 14px; }
        .panes { grid-template-columns: 1fr; }
        textarea, .editor { min-height: 220px; }
        .stats { grid-template-columns: repeat(2, 1fr); }
        .card { padding: 16px; }
        .toolbar .spacer { display: none; }
        .pane-actions .btn { flex: 1 1 auto; }
    }
    @media (max-width: 480px) {
        .stats { grid-template-columns: 1fr; }
        .nav-tab { padding: 7px 14px; font-size: 13px; }
    }
</style>
</head>
<body>

<div class="header">
    <h1>Cyrillic Word Cleaner</h1>
    <p>Replace Cyrillic &amp; lookalike Unicode characters with proper Latin text &mdash; nothing else is changed</p>
    <div class="nav-tabs">
        <a href="index.php" class="nav-tab">DOCX Generator</a>
        <a href="meta-extractor.php" class="nav-tab">Meta Extractor</a>
        <a href="meta-tool/index.php" class="nav-tab">Meta Translator</a>
        <a href="lang-generator.php" class="nav-tab">Language Tab Generator</a>
        <a href="cyrillic-cleaner.php" class="nav-tab active">Cyrillic Cleaner</a>
    </div>
</div>

<div class="main">

    <!-- Controls -->
    <div class="card">
        <div class="toolbar">
            <span class="dict-line" id="dictLine">
                <span class="dot" id="dictDot"></span>
                <span id="dictText">Loading dictionary&hellip;</span>
                <a id="dictReload" title="Re-read crilic-wordss.csv from disk">Reload CSV</a>
            </span>
        </div>

        <div class="modes">
            <span class="modes-label">Mode</span>
            <label class="mode"><input type="radio" name="mode" id="modeRich" value="rich" checked>
                Rich text <span class="hint">(keeps headings, bold, links &mdash; paste from WordPress or Word)</span></label>
            <label class="mode"><input type="radio" name="mode" id="modePlain" value="plain">
                Plain text <span class="hint">(no formatting)</span></label>
        </div>

        <div class="options">
            <label class="opt"><input type="checkbox" id="optAuto" checked> Auto Clean <span class="hint">(clean as you type or paste)</span></label>
            <label class="opt"><input type="checkbox" id="optChanges" checked> Show Changes</label>
            <label class="opt"><input type="checkbox" id="optUrls" checked> Protect URLs &amp; emails <span class="hint">(never rewrite a link)</span></label>
            <label class="opt"><input type="checkbox" id="optInvisible"> Strip invisible characters <span class="hint">(zero-width, soft hyphen, NBSP)</span></label>
            <label class="opt"><input type="checkbox" id="optTidy" checked> Clean up markup <span class="hint">(drop &lt;div&gt; and &lt;span&gt; wrappers)</span></label>
            <label class="opt"><input type="checkbox" id="optForce"> Force-convert all-Cyrillic words <span class="hint">(unsafe for real Cyrillic text)</span></label>
        </div>
    </div>

    <!-- Results -->
    <div class="card" id="resultCard">
        <div class="card-title">Results</div>

        <div class="stats">
            <div class="stat"><div class="num" id="statChars">0</div><div class="lbl">Characters Replaced</div></div>
            <div class="stat"><div class="num" id="statWords">0</div><div class="lbl">Words Corrected</div></div>
            <div class="stat"><div class="num" id="statTotal">0</div><div class="lbl">Total Replacements</div></div>
            <div class="stat" id="statRemainingBox"><div class="num" id="statRemaining">0</div><div class="lbl">Suspicious Characters Remaining</div></div>
        </div>

        <div class="banner info" id="banner">
            <span id="bannerIcon">&#8505;</span>
            <span class="body" id="bannerBody">Paste text below to begin.</span>
        </div>

        <div id="remainingWrap" class="hidden">
            <div class="codepoints" id="remainingList"></div>
        </div>
    </div>


    <!-- Input / Output -->
    <div class="card">
        <div class="panes">
            <div class="pane">
                <div class="pane-head">
                    <label for="input">Input</label>
                    <span class="meta" id="inMeta">0 characters</span>
                </div>
                <div class="editor-toolbar" id="editorToolbar">
                    <select id="fmtBlock" title="Paragraph format">
                        <option value="p">Paragraph</option>
                        <option value="h1">Heading 1</option>
                        <option value="h2">Heading 2</option>
                        <option value="h3">Heading 3</option>
                        <option value="h4">Heading 4</option>
                        <option value="blockquote">Quote</option>
                        <option value="pre">Preformatted</option>
                    </select>
                    <span class="tb-sep"></span>
                    <button type="button" class="tb-btn" data-cmd="bold" title="Bold (Ctrl+B)"><b>B</b></button>
                    <button type="button" class="tb-btn" data-cmd="italic" title="Italic (Ctrl+I)"><i>I</i></button>
                    <button type="button" class="tb-btn" data-cmd="underline" title="Underline (Ctrl+U)"><u>U</u></button>
                    <span class="tb-sep"></span>
                    <button type="button" class="tb-btn" data-cmd="insertUnorderedList" title="Bulleted list">&bull;&nbsp;List</button>
                    <button type="button" class="tb-btn" data-cmd="insertOrderedList" title="Numbered list">1.&nbsp;List</button>
                    <button type="button" class="tb-btn" data-cmd="createLink" title="Insert link">Link</button>
                    <span class="tb-sep"></span>
                    <button type="button" class="tb-btn" data-cmd="removeFormat" title="Clear formatting">Clear format</button>
                </div>
                <div class="editor" id="richInput" contenteditable="true" spellcheck="false"
                     data-placeholder="Paste or type your formatted text here..."></div>
                <textarea id="input" class="hidden" placeholder="Paste your text here..." spellcheck="false"></textarea>
                <div class="pane-actions">
                    <button type="button" class="btn btn-primary" id="btnClean">Clean Text</button>
                </div>
            </div>
            <div class="pane">
                <div class="pane-head">
                    <label for="output">Output</label>
                    <span class="meta" id="outMeta">0 characters</span>
                </div>
                <div class="editor out" id="richOutput" contenteditable="true" spellcheck="false"
                     data-placeholder="Cleaned text will appear here, with your formatting intact..."></div>
                <textarea id="output" class="hidden" placeholder="Cleaned English text will appear here..." spellcheck="false" readonly></textarea>
                <div class="pane-actions">
                    <button type="button" class="btn btn-secondary" id="btnCopy" disabled>Copy Result</button>
                    <button type="button" class="btn btn-ghost" id="btnClear">Clear</button>
                </div>
            </div>
        </div>
    </div>

    <!-- Change log -->
    <div class="card" id="changesCard">
        <div class="card-title">Change Log</div>
        <div class="table-wrap">
            <table id="changeTable">
                <thead>
                    <tr>
                        <th>Original</th>
                        <th>Corrected</th>
                        <th>Type</th>
                        <th class="num">Count</th>
                    </tr>
                </thead>
                <tbody id="changeBody"></tbody>
            </table>
        </div>
        <p class="muted" id="changeEmpty">No changes yet.</p>
    </div>

</div>

<script>
(function () {
    'use strict';

    /* =================================================================
     * Layer 2 - generic lookalike character map.
     *
     * Only characters whose Latin counterpart is unambiguous are listed.
     * Anything not in here is reported as "suspicious remaining" rather
     * than guessed at.
     *
     * This is the built-in baseline. Any character mappings found in
     * crilic-wordss.csv (rows like "а → a") are merged over the top, so the
     * CSV can extend or override it without a code change.
     * ================================================================= */
    var BUILTIN_LOOKALIKE = {
        /* --- Cyrillic --- */
        'а': 'a', 'А': 'A',   // а А
        'В': 'B',                  // В
        'е': 'e', 'Е': 'E',   // е Е
        'к': 'k', 'К': 'K',   // к К
        'М': 'M',                  // М
        'Н': 'H',                  // Н
        'о': 'o', 'О': 'O',   // о О
        'р': 'p', 'Р': 'P',   // р Р
        'с': 'c', 'С': 'C',   // с С
        'Т': 'T',                  // Т
        'у': 'y', 'У': 'Y',   // у У
        'х': 'x', 'Х': 'X',   // х Х
        'ѕ': 's', 'Ѕ': 'S',   // ѕ Ѕ
        'і': 'i', 'І': 'I',   // і І
        'ј': 'j', 'Ј': 'J',   // ј Ј
        'ӏ': 'l', 'Ӏ': 'I',   // ӏ Ӏ
        'ԁ': 'd', 'Ԁ': 'D',   // ԁ Ԁ  (U+0501 appears in crilic-wordss.csv)
        'ԛ': 'q', 'Ԛ': 'Q',   // ԛ Ԛ
        'ԝ': 'w', 'Ԝ': 'W',   // ԝ Ԝ
        'һ': 'h', 'Һ': 'H',   // һ Һ
        'ү': 'y', 'Ү': 'Y',   // ү Ү
        'ѵ': 'v', 'Ѵ': 'V',   // ѵ Ѵ

        /* --- Greek --- */
        'Α': 'A', 'Β': 'B', 'Ε': 'E', 'Ζ': 'Z', 'Η': 'H',
        'Ι': 'I', 'Κ': 'K', 'Μ': 'M', 'Ν': 'N', 'Ο': 'O',
        'Ρ': 'P', 'Τ': 'T', 'Υ': 'Y', 'Χ': 'X',
        'ο': 'o', 'ι': 'i', 'κ': 'k', 'ν': 'v',
        'ρ': 'p', 'υ': 'u', 'χ': 'x',

        /* --- Armenian --- */
        'օ': 'o', 'ո': 'n',

        /* --- Latin lookalikes from other blocks --- */
        'ɑ': 'a',  // ɑ
        'ɡ': 'g',  // ɡ

        /* --- Roman numeral forms --- */
        'Ⅰ': 'I', 'Ⅴ': 'V', 'Ⅹ': 'X', 'Ⅼ': 'L',
        'Ⅽ': 'C', 'Ⅾ': 'D', 'Ⅿ': 'M',
        'ⅰ': 'i', 'ⅴ': 'v', 'ⅹ': 'x', 'ⅼ': 'l',
        'ⅽ': 'c', 'ⅾ': 'd', 'ⅿ': 'm'
    };

    // Fullwidth Latin letters and digits are unambiguous - generate them.
    for (var i = 0; i < 26; i++) {
        BUILTIN_LOOKALIKE[String.fromCharCode(0xFF21 + i)] = String.fromCharCode(65 + i);
        BUILTIN_LOOKALIKE[String.fromCharCode(0xFF41 + i)] = String.fromCharCode(97 + i);
    }
    for (var d = 0; d < 10; d++) {
        BUILTIN_LOOKALIKE[String.fromCharCode(0xFF10 + d)] = String.fromCharCode(48 + d);
    }

    /* The map the cleaner actually uses. buildDictionary() rebuilds it as
       built-ins + whatever character mappings the CSV supplies. */
    var LOOKALIKE = BUILTIN_LOOKALIKE;

    function applyCsvChars(csvChars) {
        var merged = {}, k;
        for (k in BUILTIN_LOOKALIKE) {
            if (Object.prototype.hasOwnProperty.call(BUILTIN_LOOKALIKE, k)) merged[k] = BUILTIN_LOOKALIKE[k];
        }
        var added = 0, overrode = 0;
        for (k in (csvChars || {})) {
            if (!Object.prototype.hasOwnProperty.call(csvChars, k)) continue;
            if (k.length === 0) continue;
            if (Object.prototype.hasOwnProperty.call(merged, k)) {
                if (merged[k] !== csvChars[k]) overrode++;
            } else {
                added++;
            }
            merged[k] = csvChars[k];
        }
        LOOKALIKE = merged;
        return { added: added, overrode: overrode };
    }

    /* Invisible / formatting characters that corrupted text usually carries along.
       Built from codepoints on purpose - as raw glyphs these are literally invisible.

       Replacement ' ' = a visible space variant (harmless to leave alone).
       Replacement ''  = a truly hidden character (always worth reporting). */
    var INVISIBLE = {};        // char -> replacement
    var INVISIBLE_NAMES = {};  // char -> Unicode name, for the "remaining" report
    var HIDDEN_RE;             // matches only the truly hidden ones
    var INVISIBLE_RE;          // matches everything in INVISIBLE
    (function () {
        var spec = [
            [0x00A0, ' ', 'NO-BREAK SPACE'],
            [0x2007, ' ', 'FIGURE SPACE'],
            [0x2009, ' ', 'THIN SPACE'],
            [0x202F, ' ', 'NARROW NO-BREAK SPACE'],
            [0x205F, ' ', 'MEDIUM MATHEMATICAL SPACE'],
            [0x3000, ' ', 'IDEOGRAPHIC SPACE'],
            [0x00AD, '', 'SOFT HYPHEN'],
            [0x180E, '', 'MONGOLIAN VOWEL SEPARATOR'],
            [0x200B, '', 'ZERO WIDTH SPACE'],
            [0x200C, '', 'ZERO WIDTH NON-JOINER'],
            [0x200D, '', 'ZERO WIDTH JOINER'],
            [0x200E, '', 'LEFT-TO-RIGHT MARK'],
            [0x200F, '', 'RIGHT-TO-LEFT MARK'],
            [0x202A, '', 'LEFT-TO-RIGHT EMBEDDING'],
            [0x202B, '', 'RIGHT-TO-LEFT EMBEDDING'],
            [0x202C, '', 'POP DIRECTIONAL FORMATTING'],
            [0x202D, '', 'LEFT-TO-RIGHT OVERRIDE'],
            [0x202E, '', 'RIGHT-TO-LEFT OVERRIDE'],
            [0x2060, '', 'WORD JOINER'],
            [0x2061, '', 'FUNCTION APPLICATION'],
            [0x2062, '', 'INVISIBLE TIMES'],
            [0x2063, '', 'INVISIBLE SEPARATOR'],
            [0x2064, '', 'INVISIBLE PLUS'],
            [0xFEFF, '', 'ZERO WIDTH NO-BREAK SPACE']
        ];

        var all = '', hidden = '';
        spec.forEach(function (row) {
            var ch = String.fromCharCode(row[0]);
            INVISIBLE[ch] = row[1];
            INVISIBLE_NAMES[ch] = row[2];
            all += ch;
            if (row[1] === '') hidden += ch;
        });

        INVISIBLE_RE = new RegExp('[' + all + ']', 'g');
        HIDDEN_RE    = new RegExp('[' + hidden + ']', 'g');
    })();

    /* Scripts whose letters are commonly used as Latin lookalikes.
       Used only for *reporting* whatever is left over after cleaning. */
    var SUSPICIOUS_RE = new RegExp(
        '[' +
        '\u0370-\u03FF\u1F00-\u1FFF' +          // Greek
        '\u0400-\u04FF\u0500-\u052F' +          // Cyrillic + supplement
        '\u2DE0-\u2DFF\uA640-\uA69F' +          // Cyrillic extended
        '\u0530-\u058F' +                         // Armenian
        '\u13A0-\u13FF\uAB70-\uABBF' +          // Cherokee
        '\u2160-\u217F' +                         // Roman numeral forms
        '\uFF01-\uFF5E' +                         // Fullwidth forms
        '\u0251\u0261' +                          // IPA lookalikes
        ']', 'g'
    );

    var TOKEN_RE        = /[\p{L}\p{M}\p{N}]+/gu;
    var SINGLE_TOKEN_RE = /^[\p{L}\p{M}\p{N}]+$/u;
    var HAS_LATIN = /[A-Za-zＡ-Ｚａ-ｚ]/;

    /* =================================================================
     * Layer 1 - dictionary loaded from crilic-wordss.csv
     * ================================================================= */
    var dict = {
        loaded: false,
        exact: new Map(),        // corrupted word  -> correct word
        ci: new Map(),           // lowercased      -> correct word
        phrases: [],             // multi-word entries, longest first
        phraseRe: null,
        count: 0,
        source: 'crilic-wordss.csv'
    };

    function escapeRe(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }

    function buildDictionary(payload) {
        dict.exact = new Map();
        dict.ci = new Map();
        dict.phrases = [];
        dict.phraseRe = null;

        var map = payload.map || {};
        Object.keys(map).forEach(function (from) {
            var to = map[from];
            dict.exact.set(from, to);
            var lower = from.toLowerCase();
            if (!dict.ci.has(lower)) dict.ci.set(lower, to);
            // Anything that is not a single word token cannot be matched by the
            // token walk, so it needs the phrase pass instead.
            if (!SINGLE_TOKEN_RE.test(from)) dict.phrases.push(from);
        });

        if (dict.phrases.length) {
            // Longest first, so "delayed flight" wins over "flight".
            dict.phrases.sort(function (a, b) { return b.length - a.length; });
            var W = '[\\p{L}\\p{M}\\p{N}]';
            dict.phraseRe = new RegExp(
                '(?<!' + W + ')(?:' + dict.phrases.map(escapeRe).join('|') + ')(?!' + W + ')',
                'gu'
            );
        }

        // Layer 2: the CSV's own character list extends/overrides the built-ins.
        dict.charStats = applyCsvChars(payload.chars);
        dict.charCount = payload.charCount || 0;

        dict.count = dict.exact.size;
        dict.source = payload.source || 'crilic-wordss.csv';
        dict.loaded = true;
    }

    function loadDictionary(refresh) {
        var el = {
            dot: document.getElementById('dictDot'),
            text: document.getElementById('dictText')
        };
        el.dot.className = 'dot';
        el.text.textContent = 'Loading dictionary…';

        return fetch('cyrillic-dictionary.php' + (refresh ? '?refresh=1&_=' + Date.now() : ''), { cache: 'no-cache' })
            .then(function (r) { return r.json(); })
            .then(function (payload) {
                if (!payload || !payload.ok) {
                    throw new Error((payload && payload.error) || 'Dictionary could not be built.');
                }
                buildDictionary(payload);
                el.dot.className = 'dot ok';

                var bits = ['<b>' + payload.count + '</b> word mappings'];
                if (payload.charCount) {
                    bits.push('<b>' + payload.charCount + '</b> character mappings');
                }
                var line = bits.join(' + ') + ' from <b>' + escapeHtml(payload.source) + '</b>';

                // Say plainly when a line in the CSV could not be interpreted, so an
                // edit in an unexpected format is never silently ignored.
                if (payload.ignored && payload.ignored.length) {
                    line += ' <span class="skipped" title="' +
                        escapeHtml(payload.ignored.join('\n')) + '">&#9888; ' +
                        payload.ignored.length + ' line' +
                        (payload.ignored.length === 1 ? '' : 's') + ' not recognised</span>';
                }
                el.text.innerHTML = line;

                if (payload.warnings && payload.warnings.length) {
                    console.warn('[Cyrillic Cleaner] dictionary warnings:', payload.warnings);
                }
                if (payload.ignored && payload.ignored.length) {
                    console.warn('[Cyrillic Cleaner] lines not recognised in ' +
                        payload.source + ':', payload.ignored);
                }
            })
            .catch(function (err) {
                dict.loaded = false;
                dict.count = 0;
                LOOKALIKE = BUILTIN_LOOKALIKE;   // built-ins still clean characters
                el.dot.className = 'dot err';
                el.text.innerHTML = 'CSV dictionary unavailable — ' + escapeHtml(err.message) +
                    ' Character-level cleaning still works.';
            });
    }

    /* =================================================================
     * Protected spans - URLs, emails and HTML link attributes.
     *
     * A lookalike character inside a URL changes where the link points,
     * so those are reported instead of silently "fixed".
     * ================================================================= */
    var PROTECT_RES = [
        /\b(?:https?|ftp):\/\/[^\s<>"'`\\]+/giu,
        /\bwww\.[^\s<>"'`\\]+/giu,
        /[^\s<>"'`@(),;:\[\]]+@[^\s<>"'`@(),;:\[\]]+\.[\p{L}]{2,}/gu,
        /\b(?:href|src|srcset|action|cite|content|data-[\w-]+)\s*=\s*(?:"[^"]*"|'[^']*'|[^\s>]+)/giu
    ];

    function protectedRanges(text, enabled) {
        if (!enabled) return [];
        var ranges = [];
        PROTECT_RES.forEach(function (re) {
            re.lastIndex = 0;
            var m;
            while ((m = re.exec(text)) !== null) {
                if (m[0].length === 0) { re.lastIndex++; continue; }
                ranges.push([m.index, m.index + m[0].length]);
            }
        });
        if (!ranges.length) return ranges;
        ranges.sort(function (a, b) { return a[0] - b[0]; });
        var merged = [ranges[0]];
        for (var i = 1; i < ranges.length; i++) {
            var last = merged[merged.length - 1];
            if (ranges[i][0] <= last[1]) {
                if (ranges[i][1] > last[1]) last[1] = ranges[i][1];
            } else {
                merged.push(ranges[i]);
            }
        }
        return merged;
    }

    /* Ranges are sorted; tokens arrive in order, so a moving cursor is enough. */
    function makeRangeCursor(ranges) {
        var idx = 0;
        return function (start, end) {
            while (idx < ranges.length && ranges[idx][1] <= start) idx++;
            return idx < ranges.length && ranges[idx][0] < end;
        };
    }

    /* =================================================================
     * Case handling (section 9): only used for the case-insensitive CSV
     * fallback. Exact CSV hits keep the CSV's own capitalisation.
     * ================================================================= */
    function isUpper(ch) { return ch !== ch.toLowerCase() && ch === ch.toUpperCase(); }

    function applyCase(src, target) {
        var letters = src.replace(/[^\p{L}]/gu, '');
        if (letters.length > 1 && letters === letters.toUpperCase() && letters !== letters.toLowerCase()) {
            return target.toUpperCase();
        }
        if (src.length && isUpper(src.charAt(0))) {
            return target.charAt(0).toUpperCase() + target.slice(1);
        }
        return target;
    }

    /* =================================================================
     * The cleaner
     * ================================================================= */
    /* One accumulator can span many strings: in rich-text mode every text node
       in the pasted document is cleaned into the same set of counters. */
    function makeAcc() {
        return { words: 0, chars: 0, invisible: 0, log: new Map() };
    }

    var NON_ASCII_RE = /[^\x00-\x7F]/;

    function cleanText(text, opts, acc) {
        /* Fast path: pure ASCII cannot hold a lookalike, a CSV key (they all
           contain non-ASCII by construction) or an invisible character - so
           there is nothing to do. Most text nodes in a real document land here. */
        if (!text || !NON_ASCII_RE.test(text)) return text;

        var stats = acc;
        var log = acc.log;

        function note(from, to, type) {
            var key = from + ' ' + to + ' ' + type;
            var row = log.get(key);
            if (row) { row.count++; } else { log.set(key, { from: from, to: to, type: type, count: 1 }); }
        }

        var out = text;

        /* --- Step 3a: multi-word CSV entries (if the CSV ever gains any) --- */
        if (dict.phraseRe) {
            var phraseRanges = protectedRanges(out, opts.protectUrls);
            var phraseHit = makeRangeCursor(phraseRanges);
            dict.phraseRe.lastIndex = 0;
            out = out.replace(dict.phraseRe, function (match, offset) {
                if (phraseHit(offset, offset + match.length)) return match;
                var to = dict.exact.get(match);
                if (to === undefined) return match;
                stats.words++;
                note(match, to, 'CSV mapping');
                return to;
            });
        }

        /* --- Steps 3b + 4 + 5: token walk --- */
        var ranges = protectedRanges(out, opts.protectUrls);
        var hit = makeRangeCursor(ranges);
        var pieces = [];
        var last = 0;
        var m;

        TOKEN_RE.lastIndex = 0;
        while ((m = TOKEN_RE.exec(out)) !== null) {
            var tok = m[0];
            var start = m.index;
            var end = start + tok.length;

            if (start > last) pieces.push(out.slice(last, start));
            last = end;

            if (hit(start, end)) { pieces.push(tok); continue; }

            /* Step 3 - exact CSV word mapping */
            var mapped = dict.exact.get(tok);
            if (mapped !== undefined) {
                stats.words++;
                note(tok, mapped, 'CSV mapping');
                pieces.push(mapped);
                continue;
            }

            /* Step 3 (cont.) - same word in a different capitalisation */
            var ciHit = dict.ci.get(tok.toLowerCase());
            if (ciHit !== undefined) {
                var cased = applyCase(tok, ciHit);
                if (cased !== tok) {
                    stats.words++;
                    note(tok, cased, 'CSV mapping');
                    pieces.push(cased);
                    continue;
                }
            }

            /* Steps 4 + 5 - character-level lookalike scan */
            var hasLookalike = false;
            var j;
            for (j = 0; j < tok.length; j++) {
                if (LOOKALIKE[tok.charAt(j)] !== undefined) { hasLookalike = true; break; }
            }
            if (!hasLookalike) { pieces.push(tok); continue; }

            /* A token with no Latin letter at all is most likely genuine
               Cyrillic/Greek text, not corrupted English - leave it alone
               unless the user explicitly asks otherwise (section 15). */
            if (!HAS_LATIN.test(tok) && !opts.force) { pieces.push(tok); continue; }

            var rebuilt = '';
            for (j = 0; j < tok.length; j++) {
                var ch = tok.charAt(j);
                var rep = LOOKALIKE[ch];
                if (rep !== undefined) {
                    rebuilt += rep;
                    stats.chars++;
                    note(ch, rep, 'Unicode');
                } else {
                    rebuilt += ch;
                }
            }
            pieces.push(rebuilt);
        }
        if (last < out.length) pieces.push(out.slice(last));
        out = pieces.join('');

        /* --- Optional: invisible / formatting characters --- */
        if (opts.stripInvisible) {
            out = out.replace(INVISIBLE_RE, function (ch) {
                var rep = INVISIBLE[ch];
                if (rep === undefined) return ch;
                stats.chars++;
                stats.invisible++;
                note(codeLabel(ch), rep === '' ? '(removed)' : '(space)', 'Invisible');
                return rep;
            });
        }

        return out;
    }

    /* Steps 6 + 7: turn an accumulator into the result the UI renders.
       `scanTarget` is whatever should be searched for leftovers - the cleaned
       plain text, or the cleaned HTML (so lookalikes hiding in an href count too). */
    function summarize(acc, scanTarget) {
        var changes = Array.from(acc.log.values()).sort(function (a, b) {
            return b.count - a.count || a.from.localeCompare(b.from);
        });
        return {
            words: acc.words,
            chars: acc.chars,
            invisible: acc.invisible,
            total: acc.words + acc.chars,
            changes: changes,
            remaining: scanSuspicious(scanTarget)
        };
    }

    /* ---- Plain-text mode ---- */
    function cleanPlain(text, opts) {
        var acc = makeAcc();
        var out = cleanText(text, opts, acc);
        var result = summarize(acc, out);
        result.text = out;
        return result;
    }

    /* ---- Rich-text mode ------------------------------------------------
     * Only text nodes are touched. Tags, attributes, classes, hrefs, inline
     * styles and the document structure are left exactly as pasted, so the
     * formatting that comes in is the formatting that goes out.
     * ------------------------------------------------------------------ */
    var SKIP_TAGS = { SCRIPT: 1, STYLE: 1, NOSCRIPT: 1, TEXTAREA: 1 };

    /* ---- Markup tidy-up -------------------------------------------------
     * Editors and clipboards wrap pasted content in <div> and <span> shells
     * (<span style="font-weight:400"> and friends). They carry no meaning and
     * follow the text into WordPress, so they are unwrapped by default.
     *
     * Semantic tags - headings, p, strong/em/u, lists, links, tables - are
     * always kept, so the visible formatting is unchanged.
     * ------------------------------------------------------------------ */
    var BLOCK_TAGS = {
        P: 1, DIV: 1, H1: 1, H2: 1, H3: 1, H4: 1, H5: 1, H6: 1,
        UL: 1, OL: 1, LI: 1, BLOCKQUOTE: 1, PRE: 1, HR: 1,
        TABLE: 1, THEAD: 1, TBODY: 1, TFOOT: 1, TR: 1, TD: 1, TH: 1,
        SECTION: 1, ARTICLE: 1, ASIDE: 1, HEADER: 1, FOOTER: 1, MAIN: 1, NAV: 1,
        FIGURE: 1, FIGCAPTION: 1, DL: 1, DT: 1, DD: 1, FORM: 1
    };

    function unwrap(el) {
        var parent = el.parentNode;
        if (!parent) return;
        while (el.firstChild) parent.insertBefore(el.firstChild, el);
        parent.removeChild(el);
    }

    function hasBlockChild(el) {
        for (var i = 0; i < el.children.length; i++) {
            if (BLOCK_TAGS[el.children[i].tagName]) return true;
        }
        return false;
    }

    function tidyMarkup(root) {
        // Spans are inline - unwrapping never changes the document structure.
        var spans = root.querySelectorAll('span');
        for (var i = spans.length - 1; i >= 0; i--) unwrap(spans[i]);

        /* A <div> is either a wrapper around real blocks (unwrap it) or a
           stand-in for a paragraph (promote it to <p>, so the line survives). */
        var divs = root.querySelectorAll('div');
        for (var j = divs.length - 1; j >= 0; j--) {
            var d = divs[j];
            if (!d.parentNode) continue;

            if (hasBlockChild(d)) { unwrap(d); continue; }

            if (!d.textContent.trim() && !d.querySelector('img, br, hr')) {
                d.parentNode.removeChild(d);
                continue;
            }

            var p = document.createElement('p');
            while (d.firstChild) p.appendChild(d.firstChild);
            d.parentNode.replaceChild(p, d);
        }
        return root;
    }

    /* Pasted markup is arbitrary web content and we re-render it via innerHTML,
       so drop anything executable before it goes back into the page. */
    function sanitize(root) {
        var bad = root.querySelectorAll('script, style, noscript, iframe, object, embed, link, meta');
        for (var i = 0; i < bad.length; i++) {
            if (bad[i].parentNode) bad[i].parentNode.removeChild(bad[i]);
        }
        var all = root.querySelectorAll('*');
        for (var j = 0; j < all.length; j++) {
            var attrs = all[j].attributes;
            for (var k = attrs.length - 1; k >= 0; k--) {
                var name = attrs[k].name.toLowerCase();
                var value = (attrs[k].value || '').replace(/\s+/g, '').toLowerCase();
                if (name.indexOf('on') === 0 || value.indexOf('javascript:') === 0) {
                    all[j].removeAttribute(attrs[k].name);
                }
            }
        }
        return root;
    }

    function cleanHtml(html, opts) {
        var acc = makeAcc();
        var container = document.createElement('div');
        container.innerHTML = html;
        sanitize(container);

        var walker = document.createTreeWalker(container, NodeFilter.SHOW_TEXT, null, false);
        var nodes = [], node;
        while ((node = walker.nextNode())) nodes.push(node);

        for (var i = 0; i < nodes.length; i++) {
            var n = nodes[i];
            if (n.parentNode && SKIP_TAGS[n.parentNode.nodeName]) continue;
            var cleaned = cleanText(n.nodeValue, opts, acc);
            if (cleaned !== n.nodeValue) n.nodeValue = cleaned;
        }

        if (opts.tidy) tidyMarkup(container);

        var outHtml = container.innerHTML;
        var result = summarize(acc, outHtml);
        result.html = outHtml;
        result.text = container.textContent || '';
        return result;
    }

    function scanSuspicious(text) {
        var counts = new Map();
        var total = 0;
        var m;

        SUSPICIOUS_RE.lastIndex = 0;
        while ((m = SUSPICIOUS_RE.exec(text)) !== null) {
            counts.set(m[0], (counts.get(m[0]) || 0) + 1);
            total++;
        }
        /* Only the truly hidden characters are reported. NBSP and friends render
           as an ordinary space, so flagging them would drown out the real finds. */
        HIDDEN_RE.lastIndex = 0;
        while ((m = HIDDEN_RE.exec(text)) !== null) {
            counts.set(m[0], (counts.get(m[0]) || 0) + 1);
            total++;
        }

        var list = Array.from(counts.entries()).map(function (e) {
            return { ch: e[0], count: e[1], invisible: INVISIBLE[e[0]] !== undefined };
        }).sort(function (a, b) { return b.count - a.count; });

        return { total: total, list: list };
    }

    function codePoint(ch) {
        var cp = ch.codePointAt(0).toString(16).toUpperCase();
        while (cp.length < 4) cp = '0' + cp;
        return 'U+' + cp;
    }

    function codeLabel(ch) {
        return INVISIBLE_NAMES[ch] ? INVISIBLE_NAMES[ch] + ' (' + codePoint(ch) + ')' : codePoint(ch);
    }

    /* =================================================================
     * UI
     * ================================================================= */
    var $ = function (id) { return document.getElementById(id); };

    var elInput = $('input'), elOutput = $('output');
    var elRichIn = $('richInput'), elRichOut = $('richOutput');
    var modeRich = $('modeRich'), modePlain = $('modePlain');
    var elChars = $('statChars'), elWords = $('statWords'), elTotal = $('statTotal');
    var elRemaining = $('statRemaining'), elRemainingBox = $('statRemainingBox');
    var elBanner = $('banner'), elBannerIcon = $('bannerIcon'), elBannerBody = $('bannerBody');
    var elRemainingWrap = $('remainingWrap'), elRemainingList = $('remainingList');
    var elChangesCard = $('changesCard'), elChangeBody = $('changeBody'), elChangeEmpty = $('changeEmpty');
    var elBtnCopy = $('btnCopy');
    var elInMeta = $('inMeta'), elOutMeta = $('outMeta');

    var optAuto = $('optAuto'), optChanges = $('optChanges'), optUrls = $('optUrls'),
        optInvisible = $('optInvisible'), optForce = $('optForce'), optTidy = $('optTidy');

    var AUTO_LIMIT = 400000;  // above this, auto-clean waits for the button
    var lastResult = null;

    function escapeHtml(s) {
        return String(s).replace(/[&<>"']/g, function (c) {
            return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
        });
    }

    function currentOptions() {
        return {
            protectUrls: optUrls.checked,
            stripInvisible: optInvisible.checked,
            force: optForce.checked,
            tidy: optTidy.checked
        };
    }

    function setBanner(kind, icon, html) {
        elBanner.className = 'banner ' + kind;
        elBannerIcon.innerHTML = icon;
        elBannerBody.innerHTML = html;
    }

    /* =================================================================
     * Input formatting toolbar
     *
     * execCommand is deprecated but is still the only thing every browser
     * implements for contenteditable, and it is exactly right here: whatever
     * block the user applies (h1/h2/p/list) is preserved verbatim by the
     * cleaner, which only ever rewrites text nodes.
     * ================================================================= */
    var elToolbar = $('editorToolbar');
    var elFmtBlock = $('fmtBlock');

    function exec(cmd, value) {
        elRichIn.focus();
        try { document.execCommand(cmd, false, value || null); } catch (e) { /* unsupported */ }
        syncToolbar();
        updateMeta();
        scheduleAuto();
    }

    /* Reflect the block/inline state at the cursor back into the toolbar. */
    function syncToolbar() {
        if (!isRich()) return;

        var block = '';
        try { block = (document.queryCommandValue('formatBlock') || '').toLowerCase(); } catch (e) {}
        block = block.replace(/[<>]/g, '');
        if (block === 'div' || block === '') block = 'p';
        for (var i = 0; i < elFmtBlock.options.length; i++) {
            if (elFmtBlock.options[i].value === block) { elFmtBlock.value = block; break; }
        }

        var marks = elToolbar.querySelectorAll('.tb-btn[data-cmd]');
        for (var j = 0; j < marks.length; j++) {
            var cmd = marks[j].getAttribute('data-cmd');
            var on = false;
            try { on = document.queryCommandState(cmd); } catch (e) { on = false; }
            marks[j].classList.toggle('on', !!on);
        }
    }

    // mousedown would move focus out of the editor and drop the selection.
    elToolbar.addEventListener('mousedown', function (e) {
        if (e.target.closest('.tb-btn')) e.preventDefault();
    });

    elToolbar.addEventListener('click', function (e) {
        var btn = e.target.closest('.tb-btn');
        if (!btn) return;
        var cmd = btn.getAttribute('data-cmd');

        if (cmd === 'createLink') {
            var url = window.prompt('Link URL', 'https://');
            if (url) exec('createLink', url);
            return;
        }
        exec(cmd);
    });

    elFmtBlock.addEventListener('change', function () {
        exec('formatBlock', '<' + this.value + '>');
    });

    ['keyup', 'mouseup', 'focus'].forEach(function (ev) {
        elRichIn.addEventListener(ev, syncToolbar);
    });

    /* ---- Mode ---- */
    function isRich() { return modeRich.checked; }

    function inputText()  { return isRich() ? (elRichIn.textContent  || '') : elInput.value; }
    function outputText() { return isRich() ? (elRichOut.textContent || '') : elOutput.value; }

    function applyMode() {
        var rich = isRich();
        elToolbar.classList.toggle('hidden', !rich);
        elRichIn.classList.toggle('hidden', !rich);
        elRichOut.classList.toggle('hidden', !rich);
        elInput.classList.toggle('hidden', rich);
        elOutput.classList.toggle('hidden', rich);
        if (rich) syncToolbar();
        updateMeta();
    }

    function updateMeta() {
        elInMeta.textContent = inputText().length.toLocaleString() + ' characters';
        elOutMeta.textContent = outputText().length.toLocaleString() + ' characters';
    }

    function renderChanges(changes) {
        elChangesCard.classList.toggle('hidden', !optChanges.checked);
        if (!optChanges.checked) return;

        if (!changes || !changes.length) {
            elChangeBody.innerHTML = '';
            elChangeEmpty.classList.remove('hidden');
            return;
        }
        elChangeEmpty.classList.add('hidden');

        var tagClass = { 'CSV mapping': 'csv', 'Unicode': 'unicode', 'Invisible': 'invis' };
        var rows = changes.map(function (c) {
            return '<tr>' +
                '<td class="mono from">' + escapeHtml(c.from) + '</td>' +
                '<td class="mono to">' + escapeHtml(c.to) + '</td>' +
                '<td><span class="tag ' + (tagClass[c.type] || 'unicode') + '">' + escapeHtml(c.type) + '</span></td>' +
                '<td class="num">' + c.count + '</td>' +
                '</tr>';
        });
        elChangeBody.innerHTML = rows.join('');
    }

    function renderRemaining(remaining) {
        if (!remaining.total) {
            elRemainingWrap.classList.add('hidden');
            elRemainingList.innerHTML = '';
            return;
        }
        elRemainingWrap.classList.remove('hidden');
        elRemainingList.innerHTML = remaining.list.slice(0, 40).map(function (r) {
            var glyph = r.invisible
                ? '<b>&#9251;</b>'
                : '<b>' + escapeHtml(r.ch) + '</b>';
            var name = r.invisible && INVISIBLE_NAMES[r.ch] ? ' ' + escapeHtml(INVISIBLE_NAMES[r.ch]) : '';
            return '<span class="cp">' + glyph + ' ' + codePoint(r.ch) + name +
                   ' <span class="n">&times;' + r.count + '</span></span>';
        }).join('') + (remaining.list.length > 40
            ? '<span class="cp n">+' + (remaining.list.length - 40) + ' more</span>' : '');
    }

    function render(result, elapsedMs) {
        lastResult = result;
        updateMeta();

        elChars.textContent = result.chars.toLocaleString();
        elWords.textContent = result.words.toLocaleString();
        elTotal.textContent = result.total.toLocaleString();
        elRemaining.textContent = result.remaining.total.toLocaleString();
        elRemainingBox.classList.toggle('warn', result.remaining.total > 0);

        elBtnCopy.disabled = result.text.length === 0;

        var timing = elapsedMs !== undefined
            ? ' <span class="muted">(' + Math.max(1, Math.round(elapsedMs)) + ' ms)</span>' : '';

        if (result.remaining.total > 0) {
            setBanner('warn', '&#9888;',
                '<strong>' + result.remaining.total.toLocaleString() +
                ' suspicious Unicode character' + (result.remaining.total === 1 ? '' : 's') +
                ' remain</strong> and were not converted automatically. ' +
                'They are either genuine non-Latin text, inside a protected URL/email, or have no ' +
                'unambiguous Latin equivalent. Inspect them below.' + timing);
        } else if (result.total === 0) {
            setBanner('ok', '&#10003;',
                '<strong>Text is clean.</strong> No Cyrillic/lookalike characters detected.' + timing);
        } else {
            setBanner('ok', '&#10003;',
                '<strong>Text is clean.</strong> ' + result.total.toLocaleString() +
                ' replacement' + (result.total === 1 ? '' : 's') + ' applied — ' +
                result.words.toLocaleString() + ' word' + (result.words === 1 ? '' : 's') +
                ', ' + result.chars.toLocaleString() + ' character' + (result.chars === 1 ? '' : 's') +
                '.' + timing);
        }

        renderRemaining(result.remaining);
        renderChanges(result.changes);
    }

    function run() {
        var opts = currentOptions();
        var t0, result;

        if (isRich()) {
            var html = elRichIn.innerHTML;
            if (!inputText().trim() && !/<(img|hr|table)/i.test(html)) { resetResults(); return; }
            t0 = performance.now();
            result = cleanHtml(html, opts);
            elRichOut.innerHTML = result.html;
        } else {
            var text = elInput.value;
            if (!text) { resetResults(); return; }
            t0 = performance.now();
            result = cleanPlain(text, opts);
            elOutput.value = result.text;
        }

        render(result, performance.now() - t0);
    }

    function resetResults() {
        lastResult = null;
        elOutput.value = '';
        elRichOut.innerHTML = '';
        elChars.textContent = '0';
        elWords.textContent = '0';
        elTotal.textContent = '0';
        elRemaining.textContent = '0';
        elRemainingBox.classList.remove('warn');
        elRemainingWrap.classList.add('hidden');
        elRemainingList.innerHTML = '';
        elChangeBody.innerHTML = '';
        elChangeEmpty.classList.remove('hidden');
        elChangesCard.classList.toggle('hidden', !optChanges.checked);
        elBtnCopy.disabled = true;
        setBanner('info', '&#8505;', 'Paste text below to begin.');
        updateMeta();
    }

    /* ---- Auto Clean (debounced) ---- */
    var timer = null;
    function scheduleAuto() {
        if (!optAuto.checked) return;
        if (timer) clearTimeout(timer);
        var len = inputText().length;
        if (len > AUTO_LIMIT) {
            setBanner('info', '&#8505;',
                'Input is ' + len.toLocaleString() + ' characters — too large for Auto Clean. ' +
                'Click <strong>Clean Text</strong> to process it.');
            return;
        }
        timer = setTimeout(run, len > 50000 ? 500 : 250);
    }

    // Both editors feed the same pipeline; contenteditable fires 'input' too.
    [elInput, elRichIn].forEach(function (el) {
        el.addEventListener('input', function () {
            // Browsers leave a stray <br> behind when a contenteditable is emptied,
            // which would keep the :empty placeholder from showing.
            if (this === elRichIn && !this.textContent && this.innerHTML !== '') {
                if (/^<br\s*\/?>$/i.test(this.innerHTML.trim())) this.innerHTML = '';
            }
            updateMeta();
            scheduleAuto();
        });

        // Paste should feel immediate — skip the debounce, and cancel the one the
        // accompanying 'input' event just scheduled so the text is not cleaned twice.
        el.addEventListener('paste', function () {
            setTimeout(function () {
                if (timer) { clearTimeout(timer); timer = null; }
                updateMeta();
                if (optAuto.checked && inputText().length <= AUTO_LIMIT) run();
            }, 0);
        });
    });

    [modeRich, modePlain].forEach(function (el) {
        el.addEventListener('change', function () {
            applyMode();
            if (inputText().trim()) run(); else resetResults();
        });
    });

    $('btnClean').addEventListener('click', run);

    $('btnClear').addEventListener('click', function () {
        if (timer) clearTimeout(timer);
        elInput.value = '';
        elRichIn.innerHTML = '';
        resetResults();
        (isRich() ? elRichIn : elInput).focus();
    });

    // Captured once, so rapid double-clicks cannot leave the button stuck on "Copied!".
    var COPY_LABEL = elBtnCopy.textContent;
    var copyTimer = null;

    elBtnCopy.addEventListener('click', function () {
        var rich = isRich();
        var text = outputText();
        var html = rich ? elRichOut.innerHTML : '';
        if (!text && !html) return;

        function done() {
            if (copyTimer) clearTimeout(copyTimer);
            elBtnCopy.textContent = 'Copied!';
            elBtnCopy.classList.add('btn-copied');
            copyTimer = setTimeout(function () {
                elBtnCopy.textContent = COPY_LABEL;
                elBtnCopy.classList.remove('btn-copied');
                copyTimer = null;
            }, 1800);
        }

        // Never claim a success we did not get - tell the user to copy manually.
        function failed() {
            setBanner('err', '&#9888;',
                'Could not copy automatically. Select the output and press Ctrl+C.');
        }

        /* Rich mode puts real HTML on the clipboard, so pasting back into
           WordPress, Word or Google Docs keeps the formatting. */
        function copyRich() {
            if (navigator.clipboard && window.ClipboardItem && window.isSecureContext) {
                try {
                    var item = new window.ClipboardItem({
                        'text/html':  new Blob([html], { type: 'text/html' }),
                        'text/plain': new Blob([text], { type: 'text/plain' })
                    });
                    navigator.clipboard.write([item]).then(done, selectAndCopy);
                    return;
                } catch (e) { /* older browser - fall through */ }
            }
            selectAndCopy();
        }

        // execCommand('copy') over a selection carries the formatting with it.
        function selectAndCopy() {
            var copied = false;
            try {
                var range = document.createRange();
                range.selectNodeContents(elRichOut);
                var sel = window.getSelection();
                sel.removeAllRanges();
                sel.addRange(range);
                copied = document.execCommand('copy');
                sel.removeAllRanges();
            } catch (e) { copied = false; }
            if (copied) { done(); } else { failed(); }
        }

        function copyPlainFallback() {
            var copied = false;
            elOutput.removeAttribute('readonly');
            elOutput.select();
            try { copied = document.execCommand('copy'); } catch (e) { copied = false; }
            elOutput.setAttribute('readonly', 'readonly');
            if (window.getSelection) window.getSelection().removeAllRanges();
            if (copied) { done(); } else { failed(); }
        }

        if (rich) {
            copyRich();
        } else if (navigator.clipboard && window.isSecureContext) {
            navigator.clipboard.writeText(text).then(done, copyPlainFallback);
        } else {
            copyPlainFallback();
        }
    });

    optChanges.addEventListener('change', function () {
        elChangesCard.classList.toggle('hidden', !optChanges.checked);
        if (optChanges.checked) renderChanges(lastResult ? lastResult.changes : []);
    });

    [optUrls, optInvisible, optForce, optTidy].forEach(function (el) {
        el.addEventListener('change', function () { if (inputText().trim()) run(); });
    });

    optAuto.addEventListener('change', function () { if (optAuto.checked) scheduleAuto(); });

    $('dictReload').addEventListener('click', function (e) {
        e.preventDefault();
        loadDictionary(true).then(function () { if (inputText().trim()) run(); });
    });

    /* ---- Boot ---- */
    // Make Enter produce <p> rather than <div>, so pasted and typed blocks match.
    try { document.execCommand('defaultParagraphSeparator', false, 'p'); } catch (e) {}
    applyMode();
    resetResults();
    loadDictionary(false).then(function () { if (inputText().trim()) run(); });
})();
</script>
</body>
</html>
