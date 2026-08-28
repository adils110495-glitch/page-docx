<?php
session_start();

$settingsFile = __DIR__ . '/output/.settings.json';
$hiddenProjects = [];
if (file_exists($settingsFile)) {
    $settings = json_decode(file_get_contents($settingsFile), true) ?? [];
    $hiddenProjects = $settings['hidden_projects'] ?? [];
}
?>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Meta Tag Extractor</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }

        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
            background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%);
            min-height: 100vh;
            padding: 20px;
        }

        /* ── Nav tabs ─────────────────────────────────────────────────────── */
        .nav-tabs {
            display: flex;
            justify-content: center;
            gap: 8px;
            margin-bottom: 16px;
        }

        .nav-tab {
            padding: 8px 20px;
            border-radius: 20px;
            text-decoration: none;
            font-size: 14px;
            font-weight: 600;
            color: rgba(255,255,255,0.8);
            background: rgba(255,255,255,0.15);
            transition: background 0.2s, color 0.2s;
            border: 2px solid transparent;
        }

        .nav-tab:hover { background: rgba(255,255,255,0.25); color: white; }

        .nav-tab.active {
            background: white;
            color: #11998e;
            border-color: white;
        }

        /* ── Header ───────────────────────────────────────────────────────── */
        .header { text-align: center; color: white; margin-bottom: 20px; }
        .header h1 { font-size: 32px; margin-bottom: 8px; }
        .header p  { font-size: 16px; opacity: 0.9; margin-bottom: 14px; }

        /* ── Layout ───────────────────────────────────────────────────────── */
        .main-container {
            display: grid;
            grid-template-columns: 40% 1fr;
            gap: 20px;
            max-width: 1600px;
            margin: 0 auto;
            height: calc(100vh - 160px);
        }

        /* ── Sidebar ──────────────────────────────────────────────────────── */
        .sidebar {
            background: white;
            border-radius: 12px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            overflow: hidden;
            display: flex;
            flex-direction: column;
        }

        .sidebar-header {
            background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%);
            color: white;
            padding: 20px;
            font-weight: 600;
            font-size: 18px;
        }

        .filter-bar {
            padding: 10px 15px;
            border-bottom: 1px solid #eee;
            display: flex;
            gap: 10px;
            align-items: center;
            flex-shrink: 0;
            background: #fafafa;
        }

        .filter-bar input[type="text"] {
            flex: 1;
            padding: 6px 10px;
            border: 1px solid #e0e0e0;
            border-radius: 4px;
            font-size: 12px;
            width: auto;
        }

        .filter-bar input[type="text"]:focus {
            outline: none;
            border-color: #11998e;
        }

        .show-hidden-label {
            font-size: 12px;
            color: #666;
            display: flex;
            align-items: center;
            gap: 4px;
            white-space: nowrap;
            cursor: pointer;
            font-weight: normal;
            margin-bottom: 0;
        }

        .show-hidden-label input[type="checkbox"] {
            cursor: pointer;
            width: auto;
            height: auto;
            margin: 0;
        }

        .directory-tree {
            flex: 1;
            overflow-y: auto;
            padding: 15px;
        }

        .empty-directory {
            text-align: center;
            padding: 40px 20px;
            color: #999;
            font-style: italic;
        }

        .directory-item { padding: 10px; margin-bottom: 5px; border-radius: 6px; cursor: pointer; transition: background 0.2s; font-size: 14px; }
        .directory-item:hover { background: #f0f0f0; }

        .directory-item.folder {
            font-weight: 600;
            color: #11998e;
            display: flex;
            align-items: center;
        }

        .directory-item.folder::before { content: "📁 "; }

        .directory-item.folder .folder-name { flex: 1; cursor: pointer; }

        .directory-item.folder .folder-delete-btn { display: none; margin-left: 10px; }
        .directory-item.folder:hover .folder-delete-btn { display: inline-block; }

        .directory-item.folder.is-hidden-project {
            opacity: 0.55;
            border-left: 3px solid #aaa;
            padding-left: 7px;
        }

        .folder-hide-btn {
            display: none;
            background: #6c757d !important;
            margin-left: 5px;
        }

        .directory-item.folder:hover .folder-hide-btn { display: inline-block; }

        .directory-item.folder.is-hidden-project .folder-hide-btn {
            display: inline-block;
            background: #28a745 !important;
        }

        .directory-item.file {
            padding-left: 10px;
            color: #666;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }

        .file-info { display: flex; align-items: center; flex: 1; }

        .directory-item.file .filename::before { content: "📊 "; }

        .file-actions { display: none; gap: 8px; }
        .directory-item.file:hover .file-actions { display: flex; }

        .file-action-btn {
            padding: 4px 8px;
            font-size: 11px;
            border: none;
            border-radius: 3px;
            cursor: pointer;
            text-decoration: none;
            color: white;
            font-weight: 500;
            transition: opacity 0.2s;
        }

        .btn-download { background: #28a745; }
        .btn-remove   { background: #dc3545; }

        .file-checkbox { margin-right: 8px; cursor: pointer; width: 16px; height: 16px; }

        .folder-checkbox { margin-right: 8px; cursor: pointer; }

        .bulk-actions {
            padding: 15px;
            border-top: 1px solid #eee;
            display: flex;
            gap: 10px;
            flex-wrap: wrap;
        }

        .bulk-action-btn {
            padding: 8px 16px;
            font-size: 12px;
            border: none;
            border-radius: 5px;
            cursor: pointer;
            font-weight: 600;
            transition: opacity 0.2s;
            flex: 1;
            min-width: 100px;
        }

        .bulk-action-btn:hover   { opacity: 0.9; }
        .bulk-action-btn:disabled { opacity: 0.5; cursor: not-allowed; }

        .btn-select-all  { background: #e0e0e0; color: #333; }
        .btn-bulk-download { background: #28a745; color: white; }
        .btn-bulk-delete   { background: #dc3545; color: white; }

        /* ── Content area ─────────────────────────────────────────────────── */
        .content-area {
            background: white;
            border-radius: 12px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            padding: 30px 40px;
            display: flex;
            flex-direction: column;
        }

        #metaForm { display: flex; flex-direction: column; height: 100%; }

        .form-group { margin-bottom: 20px; flex-shrink: 0; }

        .form-group.urls-group {
            flex: 1;
            display: flex;
            flex-direction: column;
            min-height: 0;
            margin-bottom: 16px;
        }

        .urls-group .textarea-wrapper { flex: 1; display: flex; flex-direction: column; min-height: 0; }

        label {
            display: block;
            color: #333;
            font-weight: 600;
            margin-bottom: 6px;
            font-size: 13px;
        }

        .help-text { font-size: 11px; color: #777; margin-top: 5px; font-style: italic; line-height: 1.3; }

        textarea {
            width: 100%;
            padding: 12px;
            border: 2px solid #e0e0e0;
            border-radius: 6px;
            font-family: 'Courier New', monospace;
            font-size: 13px;
            resize: none;
            transition: border-color 0.3s;
            flex: 1;
            min-height: 120px;
            box-sizing: border-box;
            line-height: 1.5;
        }

        textarea::placeholder { font-size: 12px; line-height: 1.6; }
        textarea:focus { outline: none; border-color: #11998e; }

        input[type="text"] {
            width: 100%;
            padding: 10px 12px;
            border: 2px solid #e0e0e0;
            border-radius: 6px;
            font-size: 13px;
            transition: border-color 0.3s;
            box-sizing: border-box;
        }

        input[type="text"]:focus { outline: none; border-color: #11998e; }

        select {
            width: 100%;
            padding: 10px 12px;
            border: 2px solid #e0e0e0;
            border-radius: 6px;
            font-size: 13px;
            transition: border-color 0.3s;
            box-sizing: border-box;
            background: white;
            cursor: pointer;
        }

        select:focus { outline: none; border-color: #11998e; }

        .project-mode-toggle { display: flex; gap: 16px; margin-bottom: 8px; }

        .project-mode-option {
            display: flex;
            align-items: center;
            gap: 5px;
            font-size: 13px;
            font-weight: normal;
            color: #444;
            cursor: pointer;
            margin-bottom: 0;
        }

        .project-mode-option input[type="radio"] { width: auto; margin: 0; cursor: pointer; }

        button {
            background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%);
            color: white;
            border: none;
            padding: 14px 32px;
            border-radius: 6px;
            font-size: 16px;
            font-weight: 600;
            cursor: pointer;
            transition: transform 0.2s, box-shadow 0.2s;
            flex-shrink: 0;
        }

        button:hover { transform: translateY(-2px); box-shadow: 0 10px 20px rgba(17,153,142,0.4); }
        button:active { transform: translateY(0); }
        button:disabled { opacity: 0.6; cursor: not-allowed; transform: none; }

        /* ── Toast ────────────────────────────────────────────────────────── */
        .toast-container {
            position: fixed;
            top: 20px;
            right: 20px;
            z-index: 9999;
            display: flex;
            flex-direction: column;
            gap: 10px;
            max-width: 420px;
        }

        .toast {
            background: white;
            border-radius: 8px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.15);
            padding: 16px 20px;
            display: flex;
            align-items: flex-start;
            gap: 12px;
            animation: slideIn 0.3s ease-out;
            position: relative;
            overflow: hidden;
        }

        .toast::before { content: ''; position: absolute; left: 0; top: 0; bottom: 0; width: 4px; }
        .toast.toast-success::before  { background: #28a745; }
        .toast.toast-warning::before  { background: #ffc107; }
        .toast.toast-error::before    { background: #dc3545; }
        .toast.toast-processing::before { background: #17a2b8; }

        .toast-icon { font-size: 20px; flex-shrink: 0; line-height: 1; }

        .toast-content { flex: 1; font-size: 14px; color: #333; line-height: 1.5; }
        .toast-title   { font-weight: 600; margin-bottom: 4px; }
        .toast-message { font-size: 13px; color: #666; }

        .toast-close {
            background: none; border: none; color: #999; cursor: pointer;
            font-size: 18px; padding: 0; width: 20px; height: 20px;
            flex-shrink: 0; line-height: 1; transition: color 0.2s;
        }

        .toast-close:hover { color: #333; }

        @keyframes slideIn  { from { transform: translateX(420px); opacity: 0; } to { transform: translateX(0); opacity: 1; } }
        @keyframes slideOut { from { transform: translateX(0); opacity: 1; } to { transform: translateX(420px); opacity: 0; } }
        .toast.hiding { animation: slideOut 0.3s ease-out forwards; }

        @media (max-width: 1024px) {
            .main-container { grid-template-columns: 1fr; height: auto; }
            .sidebar { max-height: 300px; }
        }
    </style>
</head>
<body>
<div class="toast-container" id="toastContainer"></div>

<div class="header">
    <h1>Meta Tag Extractor</h1>
    <p>Extract meta titles and descriptions from URLs — export to CSV</p>
    <div class="nav-tabs">
        <a href="index.php" class="nav-tab">DOCX Generator</a>
        <a href="meta-extractor.php" class="nav-tab active">Meta Extractor</a>
        <a href="meta-tool/index.php" class="nav-tab">Meta Translator</a>
        <a href="lang-generator.php" class="nav-tab">Language Tab Generator</a>
        <a href="cyrillic-cleaner.php" class="nav-tab">Cyrillic Cleaner</a>
    </div>
</div>

<div class="main-container">

    <!-- ── Sidebar ──────────────────────────────────────────────────────── -->
    <div class="sidebar">
        <div class="sidebar-header">CSV Output Directory</div>
        <div class="filter-bar">
            <input type="text" id="filterInput" placeholder="Filter projects and files..." oninput="applyFilter()">
            <label class="show-hidden-label">
                <input type="checkbox" id="showHidden" onchange="applyFilter()"> Show hidden
            </label>
        </div>
        <div class="directory-tree" id="directoryTree">
            <?php
            function scanForCsv($dir, $baseDir, $hiddenProjects = []) {
                if (!is_dir($dir)) {
                    echo '<div class="empty-directory">No CSV files generated yet</div>';
                    return;
                }

                $items = scandir($dir);
                $hasContent = false;

                foreach ($items as $item) {
                    if ($item === '.' || $item === '..') continue;
                    if (substr($item, 0, 1) === '.') continue; // skip hidden files like .settings.json

                    $fullPath    = $dir . '/' . $item;
                    $relativePath = str_replace($baseDir . '/', '', $fullPath);

                    if (is_dir($fullPath)) {
                        // Only show folder if it contains CSV files
                        $csvCount = countCsvFiles($fullPath);
                        if ($csvCount === 0) continue;

                        $hasContent = true;
                        $isHiddenAttr = in_array($item, $hiddenProjects) ? 'true' : 'false';

                        echo '<div class="directory-item folder" data-project="' . htmlspecialchars($item) . '" data-hidden="' . $isHiddenAttr . '">';
                        echo '<input type="checkbox" class="folder-checkbox" onclick="event.stopPropagation(); toggleFolderFiles(this);" style="margin-right:8px;">';
                        echo '<span class="folder-name" onclick="toggleFolder(this.closest(\'.directory-item.folder\'))">' . htmlspecialchars($item) . '</span>';
                        echo '<button class="file-action-btn folder-hide-btn" onclick="event.stopPropagation(); toggleHideProject(this, \'' . htmlspecialchars(addslashes($item)) . '\')">Hide</button>';
                        echo '<button class="file-action-btn btn-remove folder-delete-btn" onclick="event.stopPropagation(); deleteFolder(\'' . htmlspecialchars('output/' . $relativePath) . '\')">Delete</button>';
                        echo '</div>';
                        echo '<div class="folder-content" style="padding-left:20px;">';
                        scanForCsv($fullPath, $baseDir, $hiddenProjects);
                        echo '</div>';
                    } elseif (strtolower(pathinfo($item, PATHINFO_EXTENSION)) === 'csv') {
                        $hasContent = true;
                        echo '<div class="directory-item file">';
                        echo '<div class="file-info">';
                        echo '<input type="checkbox" class="file-checkbox" data-file="' . htmlspecialchars('output/' . $relativePath) . '">';
                        echo '<span class="filename">' . htmlspecialchars($item) . '</span>';
                        echo '</div>';
                        echo '<div class="file-actions">';
                        echo '<a href="download-csv.php?file=' . urlencode('output/' . $relativePath) . '" class="file-action-btn btn-download">Download</a>';
                        echo '<a href="remove.php?file=' . urlencode('output/' . $relativePath) . '&redirect=meta-extractor.php" onclick="return confirm(\'Delete this file?\')" class="file-action-btn btn-remove">Remove</a>';
                        echo '</div>';
                        echo '</div>';
                    }
                }

                if (!$hasContent) {
                    echo '<div class="empty-directory">No CSV files yet</div>';
                }
            }

            function countCsvFiles($dir) {
                $count = 0;
                foreach (scandir($dir) as $item) {
                    if ($item === '.' || $item === '..') continue;
                    $path = $dir . '/' . $item;
                    if (is_file($path) && strtolower(pathinfo($item, PATHINFO_EXTENSION)) === 'csv') {
                        $count++;
                    } elseif (is_dir($path)) {
                        $count += countCsvFiles($path);
                    }
                }
                return $count;
            }

            $outputDir = __DIR__ . '/output';
            scanForCsv($outputDir, $outputDir, $hiddenProjects);
            ?>
        </div>
        <div class="bulk-actions">
            <button type="button" class="bulk-action-btn btn-select-all" onclick="toggleSelectAll()">Select All</button>
            <button type="button" class="bulk-action-btn btn-bulk-download" onclick="bulkDownload()" disabled id="bulkDownloadBtn">Download Selected</button>
            <button type="button" class="bulk-action-btn btn-bulk-delete" onclick="bulkDelete()" disabled id="bulkDeleteBtn">Delete Selected</button>
        </div>
    </div>

    <!-- ── Form ─────────────────────────────────────────────────────────── -->
    <div class="content-area">
        <form method="POST" action="meta-extractor.php" id="metaForm">

            <!-- Project / Folder -->
            <div class="form-group">
                <label>Folder (Optional)</label>
                <div class="project-mode-toggle">
                    <label class="project-mode-option">
                        <input type="radio" name="project_mode" value="new" onchange="switchProjectMode('new')"> New folder
                    </label>
                    <label class="project-mode-option">
                        <input type="radio" name="project_mode" value="existing" checked onchange="switchProjectMode('existing')"> Existing folder
                    </label>
                </div>

                <div id="projectNewInput" style="display:none;">
                    <input type="text" id="projectName" placeholder="e.g. my-website-audit" />
                </div>

                <div id="projectExistingInput">
                    <?php
                    $outputDir = __DIR__ . '/output';
                    $existingProjects = [];
                    if (is_dir($outputDir)) {
                        foreach (scandir($outputDir) as $item) {
                            if ($item === '.' || $item === '..') continue;
                            if (is_dir($outputDir . '/' . $item)) {
                                $existingProjects[] = $item;
                            }
                        }
                    }
                    ?>
                    <?php if (!empty($existingProjects)): ?>
                    <select name="project" id="projectSelect">
                        <option value="">— None (root output) —</option>
                        <?php foreach ($existingProjects as $proj): ?>
                        <option value="<?php echo htmlspecialchars($proj); ?>"><?php echo htmlspecialchars($proj); ?></option>
                        <?php endforeach; ?>
                    </select>
                    <?php else: ?>
                    <div class="help-text" style="padding:8px 0;">No existing folders found. Use "New folder" to create one.</div>
                    <?php endif; ?>
                </div>

                <div class="help-text">Save the CSV inside a named subfolder, or leave empty to save in the root output directory.</div>
            </div>

            <!-- URLs -->
            <div class="form-group urls-group">
                <label for="urls">Website URLs</label>
                <div class="textarea-wrapper">
                    <textarea
                        name="urls"
                        id="urls"
                        placeholder="https://example.com/page1&#10;https://example.com/page2&#10;https://example.com/page3&#10;...&#10;One URL per line (up to 500 URLs)"
                        required
                    ></textarea>
                </div>
                <div class="help-text">Enter one URL per line (http:// or https://). Up to 500 URLs per batch.</div>
            </div>

            <div class="button-group">
                <button type="submit" id="submitBtn">Extract &amp; Download CSV</button>
            </div>
        </form>

        <!-- Progress section — visible only during batch processing -->
        <div id="progressSection" style="display:none; margin-top:20px; padding:20px; background:#f0faf8; border-radius:8px; border:1px solid #b2dfdb;">
            <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:14px;">
                <div>
                    <div style="font-size:15px; font-weight:700; color:#0d5c52;" id="progressTitle">Extracting Meta Tags…</div>
                    <div style="font-size:12px; color:#555; margin-top:3px;" id="batchStatus">Initializing…</div>
                </div>
                <button type="button" onclick="cancelProcessing()"
                        style="background:#dc3545; padding:7px 14px; font-size:12px; border-radius:4px; flex-shrink:0;">
                    Cancel
                </button>
            </div>
            <div style="background:#c8e6c9; border-radius:10px; height:22px; overflow:hidden; margin-bottom:8px;">
                <div id="progressFill"
                     style="height:100%; background:linear-gradient(135deg,#11998e 0%,#38ef7d 100%);
                            transition:width 0.4s ease; width:0%;
                            display:flex; align-items:center; justify-content:center; min-width:28px;">
                    <span id="progressPct" style="font-size:10px; font-weight:700; color:white; white-space:nowrap;"></span>
                </div>
            </div>
            <div style="font-size:12px; color:#444; margin-top:4px;" id="progressStats">0 / 0 URLs processed</div>
            <div style="font-size:12px; color:#e67e22; margin-top:3px; min-height:16px;" id="retryStats"></div>
        </div>

        <?php
        if (isset($_SESSION['meta_status'])) {
            $status = $_SESSION['meta_status'];
            echo '<div id="statusData" style="display:none;"
                      data-type="'    . htmlspecialchars($status['type'])    . '"
                      data-message="' . htmlspecialchars($status['message']) . '"';
            if (!empty($status['csv_file'])) {
                echo ' data-csv-file="' . htmlspecialchars($status['csv_file']) . '"';
            }
            echo '></div>';
            unset($_SESSION['meta_status']);
        }
        ?>
    </div>
</div>

<script>
    // ── Constants ────────────────────────────────────────────────────────
    const BATCH_SIZE  = 10;   // URLs processed in parallel per request
    const MAX_URLS    = 500;
    const MAX_RETRIES = 3;

    let isCancelled = false;

    // ── Project mode toggle ──────────────────────────────────────────────
    function switchProjectMode(mode) {
        const newInput      = document.getElementById('projectNewInput');
        const existingInput = document.getElementById('projectExistingInput');
        const nameField     = document.getElementById('projectName');
        const selectField   = document.getElementById('projectSelect');

        if (mode === 'existing') {
            newInput.style.display      = 'none';
            existingInput.style.display = '';
            if (nameField)   nameField.removeAttribute('name');
            if (selectField) selectField.setAttribute('name', 'project');
        } else {
            newInput.style.display      = '';
            existingInput.style.display = 'none';
            if (nameField)   nameField.setAttribute('name', 'project');
            if (selectField) selectField.removeAttribute('name');
        }
    }

    function getSelectedProject() {
        const mode = document.querySelector('input[name="project_mode"]:checked')?.value;
        if (mode === 'new') {
            return document.getElementById('projectName')?.value?.trim() || '';
        }
        return document.getElementById('projectSelect')?.value?.trim() || '';
    }

    // ── Hidden projects ──────────────────────────────────────────────────
    let hiddenProjects = <?php echo json_encode($hiddenProjects); ?>;

    function toggleHideProject(button, projectName) {
        const folderItem        = button.closest('.directory-item.folder');
        const isCurrentlyHidden = folderItem.dataset.hidden === 'true';

        if (isCurrentlyHidden) {
            hiddenProjects = hiddenProjects.filter(p => p !== projectName);
            folderItem.dataset.hidden = 'false';
        } else {
            if (!hiddenProjects.includes(projectName)) hiddenProjects.push(projectName);
            folderItem.dataset.hidden = 'true';
        }

        saveSettings();
        applyFilter();
    }

    function saveSettings() {
        fetch('settings.php', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ action: 'save_settings', hidden_projects: hiddenProjects })
        }).catch(() => {});
    }

    // ── Filter ───────────────────────────────────────────────────────────
    function applyFilter() {
        const filter     = document.getElementById('filterInput').value.toLowerCase().trim();
        const showHidden = document.getElementById('showHidden').checked;

        document.querySelectorAll('#directoryTree .directory-item.folder').forEach(folder => {
            const folderContent    = folder.nextElementSibling;
            const hasFolderContent = folderContent && folderContent.classList.contains('folder-content');
            const isHidden         = folder.dataset.hidden === 'true';
            const projectName      = (folder.dataset.project || '').toLowerCase();

            const hideBtn = folder.querySelector('.folder-hide-btn');
            if (hideBtn) hideBtn.textContent = isHidden ? 'Unhide' : 'Hide';

            if (isHidden && showHidden) {
                folder.classList.add('is-hidden-project');
            } else {
                folder.classList.remove('is-hidden-project');
            }

            if (isHidden && !showHidden) {
                folder.style.display = 'none';
                if (hasFolderContent) folderContent.style.display = 'none';
                return;
            }

            if (filter) {
                const folderMatches  = projectName.includes(filter);
                let hasMatchingFiles = false;

                if (hasFolderContent) {
                    folderContent.querySelectorAll('.directory-item.file').forEach(fileItem => {
                        const fileName   = (fileItem.querySelector('.filename')?.textContent || '').toLowerCase();
                        const fileMatches = folderMatches || fileName.includes(filter);
                        fileItem.style.display = fileMatches ? '' : 'none';
                        if (fileName.includes(filter)) hasMatchingFiles = true;
                    });
                }

                const show = folderMatches || hasMatchingFiles;
                folder.style.display = show ? '' : 'none';
                if (hasFolderContent) folderContent.style.display = show ? 'block' : 'none';
            } else {
                folder.style.display = '';
                if (hasFolderContent) {
                    folderContent.querySelectorAll('.directory-item.file').forEach(fi => fi.style.display = '');
                }
            }
        });

        document.querySelectorAll('#directoryTree > .directory-item.file').forEach(fileItem => {
            if (filter) {
                const fileName = (fileItem.querySelector('.filename')?.textContent || '').toLowerCase();
                fileItem.style.display = fileName.includes(filter) ? '' : 'none';
            } else {
                fileItem.style.display = '';
            }
        });
    }

    // ── Folder / file toggles ────────────────────────────────────────────
    function toggleFolder(element) {
        const folderContent = element.nextElementSibling;
        if (folderContent && folderContent.classList.contains('folder-content')) {
            folderContent.style.display = folderContent.style.display === 'none' ? 'block' : 'none';
        }
    }

    function toggleFolderFiles(folderCheckbox) {
        const folderItem    = folderCheckbox.closest('.directory-item.folder');
        const folderContent = folderItem.nextElementSibling;
        if (folderContent && folderContent.classList.contains('folder-content')) {
            folderContent.querySelectorAll('.file-checkbox').forEach(cb => {
                cb.checked = folderCheckbox.checked;
            });
            updateBulkActionButtons();
        }
    }

    // ── Bulk actions ─────────────────────────────────────────────────────
    function updateBulkActionButtons() {
        const checked = document.querySelectorAll('.file-checkbox:checked').length;
        document.getElementById('bulkDownloadBtn').disabled = checked === 0;
        document.getElementById('bulkDeleteBtn').disabled   = checked === 0;
    }

    function toggleSelectAll() {
        const checkboxes = document.querySelectorAll('.file-checkbox');
        const allChecked = Array.from(checkboxes).every(cb => cb.checked);
        checkboxes.forEach(cb => cb.checked = !allChecked);
        updateBulkActionButtons();
    }

    function bulkDownload() {
        const checkedBoxes = document.querySelectorAll('.file-checkbox:checked');
        if (!checkedBoxes.length) return;

        checkedBoxes.forEach((checkbox, index) => {
            setTimeout(() => {
                const iframe = document.createElement('iframe');
                iframe.style.display = 'none';
                iframe.src = 'download-csv.php?file=' + encodeURIComponent(checkbox.dataset.file);
                document.body.appendChild(iframe);
                setTimeout(() => document.body.removeChild(iframe), 2000);
            }, index * 500);
        });

        showToast('success', `Downloading ${checkedBoxes.length} file(s)…`);
    }

    function bulkDelete() {
        const checkedBoxes = document.querySelectorAll('.file-checkbox:checked');
        if (!checkedBoxes.length) return;
        if (!confirm(`Delete ${checkedBoxes.length} file(s)?`)) return;

        const files = Array.from(checkedBoxes).map(cb => cb.dataset.file);
        fetch('bulk_actions.php', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ action: 'delete', files })
        })
        .then(r => r.json())
        .then(data => {
            if (data.success) {
                showToast('success', `Deleted ${data.deleted} file(s)`);
                setTimeout(() => location.reload(), 1500);
            } else {
                showToast('error', data.message || 'Failed to delete files');
            }
        })
        .catch(() => showToast('error', 'An error occurred while deleting files'));
    }

    function deleteFolder(folderPath) {
        if (!confirm(`Delete the folder "${folderPath}" and all its contents?`)) return;

        fetch('bulk_actions.php', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ action: 'delete_folder', folder: folderPath })
        })
        .then(r => r.json())
        .then(data => {
            if (data.success) {
                showToast('success', `Deleted folder and ${data.deleted} file(s)`);
                setTimeout(() => location.reload(), 1500);
            } else {
                showToast('error', data.message || 'Failed to delete folder');
            }
        })
        .catch(() => showToast('error', 'An error occurred while deleting folder'));
    }

    // ── Toast ────────────────────────────────────────────────────────────
    function showToast(type, message, extra = '') {
        const container = document.getElementById('toastContainer');
        const toast     = document.createElement('div');
        toast.className = `toast toast-${type}`;

        const icons  = { success: '✓', warning: '⚠', error: '✕', processing: 'ℹ' };
        const titles = { success: 'Success', warning: 'Warning', error: 'Error', processing: 'Processing' };

        toast.innerHTML = `
            <span class="toast-icon">${icons[type] || 'ℹ'}</span>
            <div class="toast-content">
                <div class="toast-title">${titles[type] || 'Info'}</div>
                <div class="toast-message">${message}</div>
                ${extra}
            </div>
            <button class="toast-close" onclick="closeToast(this)">×</button>
        `;

        container.appendChild(toast);
        const duration = type === 'error' ? 10000 : (type === 'warning' ? 8000 : 6000);
        setTimeout(() => closeToast(toast.querySelector('.toast-close')), duration);
    }

    function closeToast(button) {
        const toast = button.closest('.toast');
        toast.classList.add('hiding');
        setTimeout(() => toast.remove(), 300);
    }

    // ── Batch processing ─────────────────────────────────────────────────
    function showProgressUI() {
        document.getElementById('metaForm').style.display        = 'none';
        document.getElementById('progressSection').style.display = '';
        document.getElementById('progressTitle').textContent     = 'Extracting Meta Tags…';
        document.getElementById('batchStatus').textContent       = 'Starting…';
        document.getElementById('progressFill').style.width      = '0%';
        document.getElementById('progressPct').textContent       = '0%';
        document.getElementById('progressStats').textContent     = '';
        document.getElementById('retryStats').textContent        = '';
    }

    function hideProgressUI() {
        document.getElementById('progressSection').style.display = 'none';
        document.getElementById('metaForm').style.display        = '';
        const btn = document.getElementById('submitBtn');
        btn.disabled    = false;
        btn.textContent = 'Extract & Download CSV';
    }

    function updateProgress(processed, total, errors, batchNum, totalBatches, retried) {
        const pct = total > 0 ? Math.round((processed / total) * 100) : 0;
        document.getElementById('progressFill').style.width  = pct + '%';
        document.getElementById('progressPct').textContent   = pct + '%';
        document.getElementById('batchStatus').textContent   = `Batch ${batchNum} of ${totalBatches}`;
        document.getElementById('progressStats').textContent =
            `${processed} / ${total} URLs processed` + (errors > 0 ? ` — ${errors} error(s)` : '');
        if (retried > 0) {
            document.getElementById('retryStats').textContent = `↻ ${retried} URL(s) needed retry`;
        }
    }

    function cancelProcessing() {
        isCancelled = true;
        hideProgressUI();
        showToast('warning', 'Processing cancelled.');
    }

    async function processBatch(batchUrls) {
        const resp = await fetch('fetch-meta.php', {
            method:  'POST',
            headers: { 'Content-Type': 'application/json' },
            body:    JSON.stringify({ action: 'process', urls: batchUrls, max_retries: MAX_RETRIES }),
        });
        if (!resp.ok) throw new Error(`Server error ${resp.status}`);
        return resp.json();
    }

    async function finalize(allResults, project) {
        const resp = await fetch('fetch-meta.php', {
            method:  'POST',
            headers: { 'Content-Type': 'application/json' },
            body:    JSON.stringify({ action: 'finalize', results: allResults, project }),
        });
        if (!resp.ok) throw new Error(`Server error ${resp.status}`);
        return resp.json();
    }

    async function startBatchProcessing(urls, project) {
        isCancelled = false;
        showProgressUI();

        // Split into batches of BATCH_SIZE
        const batches = [];
        for (let i = 0; i < urls.length; i += BATCH_SIZE) {
            batches.push(urls.slice(i, i + BATCH_SIZE));
        }

        const allResults = [];
        let processed    = 0;
        let errors       = 0;
        let retried      = 0;

        for (let i = 0; i < batches.length; i++) {
            if (isCancelled) return;

            updateProgress(processed, urls.length, errors, i + 1, batches.length, retried);

            try {
                const data = await processBatch(batches[i]);

                if (data.success && Array.isArray(data.results)) {
                    data.results.forEach(r => {
                        allResults.push(r);
                        processed++;
                        if (r.status === 'error')    errors++;
                        if ((r.attempts || 0) > 1)   retried++;
                    });
                } else {
                    // Entire batch failed at the server level
                    batches[i].forEach(url => {
                        allResults.push({ url, meta_title: 'Error: Server error', meta_description: '', status: 'error', attempts: 0 });
                        processed++;
                        errors++;
                    });
                }
            } catch (err) {
                // Network / parse error — mark batch as failed
                batches[i].forEach(url => {
                    allResults.push({ url, meta_title: 'Error: ' + err.message, meta_description: '', status: 'error', attempts: 0 });
                    processed++;
                    errors++;
                });
            }

            updateProgress(processed, urls.length, errors, i + 1, batches.length, retried);
        }

        if (isCancelled) return;

        // Write CSV
        document.getElementById('progressTitle').textContent = 'Saving CSV…';
        document.getElementById('batchStatus').textContent   = 'Writing file…';

        try {
            const result = await finalize(allResults, project);

            if (result.success) {
                const summary = `Processed ${processed} URL(s): ${processed - errors} successful, ${errors} failed`
                    + (retried > 0 ? `, ${retried} retried` : '') + '.';

                const downloadLink = `<div style="margin-top:8px;">
                    <a href="download-csv.php?file=${encodeURIComponent(result.csv_file)}"
                       style="color:#11998e;text-decoration:underline;font-weight:600;">⬇ Download CSV</a>
                </div>`;

                hideProgressUI();
                const toastType = (errors === processed) ? 'error' : (errors > 0 ? 'warning' : 'success');
                showToast(toastType, summary, downloadLink);

                // Auto-trigger download
                const iframe = document.createElement('iframe');
                iframe.style.display = 'none';
                iframe.src = 'download-csv.php?file=' + encodeURIComponent(result.csv_file);
                document.body.appendChild(iframe);
                setTimeout(() => document.body.removeChild(iframe), 4000);

                // Reload to refresh the sidebar
                setTimeout(() => location.reload(), 4500);

            } else {
                hideProgressUI();
                showToast('error', 'Failed to save CSV: ' + (result.message || 'Unknown error'));
            }
        } catch (err) {
            hideProgressUI();
            showToast('error', 'Failed to save CSV: ' + err.message);
        }
    }

    // ── Form submit intercept ────────────────────────────────────────────
    document.getElementById('metaForm').addEventListener('submit', function (e) {
        e.preventDefault();

        let urls = document.getElementById('urls').value
            .split('\n')
            .map(u => u.trim())
            .filter(u => u !== '');

        if (urls.length === 0) {
            showToast('error', 'Please enter at least one URL.');
            return;
        }

        if (urls.length > MAX_URLS) {
            showToast('warning', `Maximum ${MAX_URLS} URLs allowed. Only the first ${MAX_URLS} will be processed.`);
            urls = urls.slice(0, MAX_URLS);
        }

        const project = getSelectedProject();
        startBatchProcessing(urls, project);
    });

    // ── Init ─────────────────────────────────────────────────────────────
    document.addEventListener('DOMContentLoaded', function () {
        document.querySelectorAll('.file-checkbox').forEach(cb => {
            cb.addEventListener('change', updateBulkActionButtons);
        });
        applyFilter();
    });
</script>
</body>
</html>
