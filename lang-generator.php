<?php
session_start();

$settingsFile = __DIR__ . '/output/.settings.json';
$hiddenProjects = [];
if (file_exists($settingsFile)) {
    $settings = json_decode(file_get_contents($settingsFile), true) ?? [];
    $hiddenProjects = $settings['hidden_projects'] ?? [];
}

// Read a single key from the root .env file
function readEnvKey(string $key): string {
    $file = __DIR__ . '/.env';
    if (!file_exists($file)) return '';
    foreach (file($file, FILE_IGNORE_NEW_LINES | FILE_SKIP_EMPTY_LINES) as $line) {
        if ($line === '' || $line[0] === '#') continue;
        [$k, $v] = array_pad(explode('=', $line, 2), 2, '');
        if (trim($k) === $key) return trim($v);
    }
    return '';
}

// Google Docs auth status
$googleAuthenticated  = false;
$googleHasCredentials = false;
$googleDriveFolderId  = readEnvKey('GOOGLE_DRIVE_FOLDER_ID');
if (file_exists(__DIR__ . '/vendor/autoload.php')) {
    require_once __DIR__ . '/vendor/autoload.php';
    if (class_exists('App\GoogleAuthHelper')) {
        try {
            $googleAuth           = new \App\GoogleAuthHelper();
            $googleHasCredentials = $googleAuth->hasCredentials();
            $googleAuthenticated  = $googleAuth->isAuthenticated();
        } catch (\Exception $e) {
            // Google SDK not ready
        }
    }
}
?>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Language Tab Generator</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }

        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
        }

        .header { text-align: center; color: white; margin-bottom: 20px; }
        .header h1 { font-size: 32px; margin-bottom: 8px; }
        .header p { font-size: 16px; opacity: 0.9; }

        .main-container {
            display: grid;
            grid-template-columns: 40% 1fr;
            gap: 20px;
            max-width: 1600px;
            margin: 0 auto;
            height: calc(100vh - 140px);
        }

        .sidebar {
            background: white;
            border-radius: 12px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            overflow: hidden;
            display: flex;
            flex-direction: column;
        }

        .sidebar-header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 20px;
            font-weight: 600;
            font-size: 18px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }

        .sidebar-header .title { flex: 1; }

        .logs-link {
            color: white;
            text-decoration: none;
            font-size: 14px;
            font-weight: 500;
            padding: 6px 12px;
            background: rgba(255,255,255,0.2);
            border-radius: 4px;
            transition: background 0.2s;
        }

        .logs-link:hover { background: rgba(255,255,255,0.3); }

        .directory-tree { flex: 1; overflow-y: auto; padding: 15px; }

        .directory-item {
            padding: 10px;
            margin-bottom: 5px;
            border-radius: 6px;
            cursor: pointer;
            transition: background 0.2s;
            font-size: 14px;
        }

        .directory-item:hover { background: #f0f0f0; }

        .directory-item.folder {
            font-weight: 600;
            color: #667eea;
            display: flex;
            align-items: center;
        }

        .directory-item.folder::before { content: "📁 "; }

        .directory-item.folder .folder-name { flex: 1; cursor: pointer; }
        .directory-item.folder .folder-delete-btn { display: none; margin-left: 10px; }
        .directory-item.folder:hover .folder-delete-btn { display: inline-block; }

        .directory-item.file {
            padding-left: 10px;
            color: #666;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }

        .directory-item.file .filename::before { content: "📄 "; }
        .directory-item.log .filename::before { content: "📋 "; }

        .file-info { display: flex; align-items: center; flex: 1; }

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
        .btn-view { background: #007bff; }
        .btn-remove { background: #dc3545; }

        .empty-directory { text-align: center; padding: 40px 20px; color: #999; font-style: italic; }

        .bulk-actions {
            padding: 15px;
            border-top: 1px solid rgba(255,255,255,0.2);
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

        .bulk-action-btn:hover { opacity: 0.9; }
        .bulk-action-btn:disabled { opacity: 0.5; cursor: not-allowed; }

        .btn-bulk-download { background: #28a745; color: white; }
        .btn-bulk-delete { background: #dc3545; color: white; }
        .btn-select-all { background: rgba(255,255,255,0.2); color: white; border: 1px solid rgba(255,255,255,0.3); }

        .file-checkbox { margin-right: 8px; cursor: pointer; width: 16px; height: 16px; }

        .content-area {
            background: white;
            border-radius: 12px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            padding: 30px 40px;
            display: flex;
            flex-direction: column;
            overflow-y: auto;
        }

        .form-wrapper { display: flex; flex-direction: column; }

        #langForm { display: flex; flex-direction: column; }

        .form-group { margin-bottom: 20px; flex-shrink: 0; }

        .form-group.urls-group {
            display: flex;
            flex-direction: column;
            margin-bottom: 16px;
        }

        .urls-group .textarea-wrapper { display: flex; flex-direction: column; }

        label { display: block; color: #333; font-weight: 600; margin-bottom: 6px; font-size: 13px; }

        .help-text { font-size: 11px; color: #777; margin-top: 5px; font-style: italic; line-height: 1.3; }

        textarea {
            width: 100%;
            padding: 12px;
            border: 2px solid #e0e0e0;
            border-radius: 6px;
            font-family: 'Courier New', monospace;
            font-size: 13px;
            resize: vertical;
            transition: border-color 0.3s;
            min-height: 160px;
            height: 160px;
            box-sizing: border-box;
            line-height: 1.5;
        }

        textarea:focus { outline: none; border-color: #667eea; }

        input[type="text"] {
            width: 100%;
            padding: 10px 12px;
            border: 2px solid #e0e0e0;
            border-radius: 6px;
            font-size: 13px;
            transition: border-color 0.3s;
            box-sizing: border-box;
        }

        input[type="text"]:focus { outline: none; border-color: #667eea; }

        button {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
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

        button:hover { transform: translateY(-2px); box-shadow: 0 10px 20px rgba(102,126,234,0.4); }
        button:active { transform: translateY(0); }
        button:disabled { opacity: 0.6; cursor: not-allowed; transform: none; }

        .button-group { flex-shrink: 0; }

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

        select:focus { outline: none; border-color: #667eea; }

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

        /* Filter Bar */
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
            margin-bottom: 0;
        }

        .folder-hide-btn { display: none; background: #6c757d !important; margin-left: 5px; }
        .directory-item.folder:hover .folder-hide-btn { display: inline-block; }
        .directory-item.folder.is-hidden-project { opacity: 0.55; border-left: 3px solid #aaa; padding-left: 7px; }
        .directory-item.folder.is-hidden-project .folder-hide-btn { display: inline-block; background: #28a745 !important; }

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

        .show-hidden-label input[type="checkbox"] { cursor: pointer; width: auto; height: auto; margin: 0; }

        /* Toast */
        .toast-container {
            position: fixed;
            top: 20px;
            right: 20px;
            z-index: 9999;
            display: flex;
            flex-direction: column;
            gap: 10px;
            max-width: 400px;
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
        .toast.toast-success::before { background: #28a745; }
        .toast.toast-error::before { background: #dc3545; }
        .toast.toast-processing::before { background: #17a2b8; }
        .toast-icon { font-size: 20px; flex-shrink: 0; line-height: 1; }
        .toast-content { flex: 1; font-size: 14px; color: #333; line-height: 1.5; }
        .toast-title { font-weight: 600; margin-bottom: 4px; }
        .toast-message { font-size: 13px; color: #666; }
        .toast-close { background: none; border: none; color: #999; cursor: pointer; font-size: 18px; padding: 0; width: 20px; height: 20px; flex-shrink: 0; line-height: 1; transition: color 0.2s; }
        .toast-close:hover { color: #333; }
        .toast-progress { margin-top: 8px; height: 4px; background: #e0e0e0; border-radius: 2px; overflow: hidden; }
        .toast-progress-bar { height: 100%; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); transition: width 0.3s; }

        @keyframes slideIn { from { transform: translateX(400px); opacity: 0; } to { transform: translateX(0); opacity: 1; } }
        @keyframes slideOut { from { transform: translateX(0); opacity: 1; } to { transform: translateX(400px); opacity: 0; } }
        .toast.hiding { animation: slideOut 0.3s ease-out forwards; }

        /* Language Tabs */
        .lang-tabs-section {
            margin-bottom: 16px;
            flex-shrink: 0;
            border: 2px solid #e8ebff;
            border-radius: 10px;
            padding: 16px;
            background: #f8f9ff;
        }

        .lang-tabs-section > label {
            color: #667eea;
            font-size: 13px;
            margin-bottom: 10px;
        }

        .lang-tabs-nav {
            display: flex;
            gap: 8px;
            flex-wrap: wrap;
            margin-bottom: 12px;
        }

        .lang-tab-btn {
            padding: 6px 14px;
            border: 2px solid #667eea;
            border-radius: 20px;
            background: white;
            color: #667eea;
            font-size: 12px;
            font-weight: 700;
            cursor: pointer;
            transition: all 0.2s;
            display: inline-flex;
            align-items: center;
            gap: 5px;
            letter-spacing: 0.3px;
            box-shadow: none;
            transform: none;
        }

        .lang-tab-btn.active {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            border-color: transparent;
            box-shadow: 0 4px 12px rgba(102,126,234,0.35);
        }

        .lang-tab-btn:hover:not(.active) {
            background: #f0f2ff;
            transform: translateY(-1px);
        }

        .lang-tab-badge {
            background: rgba(102,126,234,0.15);
            border-radius: 10px;
            padding: 1px 7px;
            font-size: 11px;
            font-weight: 600;
        }

        .lang-tab-btn.active .lang-tab-badge {
            background: rgba(255,255,255,0.25);
        }

        .lang-tab-content { display: none; }
        .lang-tab-content.active { display: block; }

        .lang-urls-panel {
            background: white;
            border: 1px solid #dde1ff;
            border-radius: 6px;
            padding: 10px 12px;
            max-height: 130px;
            overflow-y: auto;
        }

        .lang-url-item {
            font-family: 'Courier New', monospace;
            font-size: 11px;
            color: #555;
            padding: 3px 0;
            border-bottom: 1px solid #f0f2ff;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }

        .lang-url-item:last-child { border-bottom: none; }

        .lang-empty { font-size: 12px; color: #aaa; font-style: italic; padding: 8px 0; }

        .lang-summary {
            font-size: 12px;
            color: #888;
            margin-bottom: 10px;
        }

        .lang-summary strong { color: #667eea; }

        @media (max-width: 1024px) {
            .main-container { grid-template-columns: 1fr; height: auto; }
            .sidebar { max-height: 300px; }
        }

        /* ── Output Destination Toggle ───────────────────────────────── */
        .output-dest-section {
            margin-bottom: 20px;
            flex-shrink: 0;
        }

        .output-dest-section > label {
            display: block;
            color: #333;
            font-weight: 600;
            margin-bottom: 10px;
            font-size: 13px;
        }

        .dest-toggle {
            display: flex;
            gap: 10px;
            margin-bottom: 14px;
        }

        .dest-option {
            flex: 1;
            position: relative;
        }

        .dest-option input[type="radio"] {
            position: absolute;
            opacity: 0;
            width: 0;
            height: 0;
        }

        .dest-option label {
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 8px;
            padding: 12px 16px;
            border: 2px solid #e0e0e0;
            border-radius: 8px;
            cursor: pointer;
            font-size: 13px;
            font-weight: 600;
            color: #555;
            background: #fafafa;
            transition: all 0.2s;
            margin-bottom: 0;
            user-select: none;
        }

        .dest-option input[type="radio"]:checked + label {
            border-color: #667eea;
            background: linear-gradient(135deg, #667eea15 0%, #764ba215 100%);
            color: #667eea;
        }

        .dest-option label:hover {
            border-color: #aaa;
            background: #f5f5f5;
        }

        .dest-option input[type="radio"]:checked + label:hover {
            background: linear-gradient(135deg, #667eea20 0%, #764ba220 100%);
        }

        .dest-icon { font-size: 18px; line-height: 1; }

        /* Google auth badge (used inside the gdrive panel) */
        .google-auth-badge {
            display: inline-flex;
            align-items: center;
            gap: 6px;
            padding: 4px 10px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 600;
        }

        .google-auth-badge.connected    { background: #e6f4ea; color: #1e8e3e; border: 1px solid #c3e6cb; }
        .google-auth-badge.disconnected { background: #fce8e6; color: #c5221f; border: 1px solid #f5c6c5; }
        .google-auth-badge.no-creds     { background: #fef7e0; color: #e37400; border: 1px solid #f8dda0; }

        .google-auth-badge .badge-dot { width: 7px; height: 7px; border-radius: 50%; flex-shrink: 0; }
        .google-auth-badge.connected .badge-dot    { background: #1e8e3e; }
        .google-auth-badge.disconnected .badge-dot { background: #c5221f; }
        .google-auth-badge.no-creds .badge-dot     { background: #e37400; }

        .google-auth-link { font-size: 12px; color: #1a73e8; text-decoration: none; font-weight: 500; }
        .google-auth-link:hover { text-decoration: underline; }

        .google-docs-notice {
            font-size: 12px; color: #666; line-height: 1.4;
            padding: 8px 10px; background: #fff3cd;
            border-radius: 5px; border-left: 3px solid #e37400;
            margin-top: 8px;
        }

        .gdrive-panel { display: none; }
        .gdrive-panel.active { display: block; }

        .gdrive-dest-info {
            display: flex;
            align-items: center;
            gap: 8px;
            padding: 8px 12px;
            background: #f0f4ff;
            border-radius: 6px;
            border: 1px solid #dde1ff;
            font-size: 12px;
            color: #444;
            margin-top: 10px;
        }

        .gdrive-dest-info .dest-folder-icon { font-size: 16px; }
        .gdrive-dest-info a { color: #1a73e8; text-decoration: none; font-weight: 500; }
        .gdrive-dest-info a:hover { text-decoration: underline; }
    </style>
</head>
<body>
    <div class="toast-container" id="toastContainer"></div>

    <div class="header">
        <h1>Website to DOCX Generator</h1>
        <p>Convert website pages into formatted DOCX documents</p>
        <div style="display:flex;justify-content:center;gap:8px;margin-top:14px;flex-wrap:wrap;">
            <a href="index.php" style="padding:8px 20px;border-radius:20px;text-decoration:none;font-size:14px;font-weight:600;color:rgba(255,255,255,0.8);background:rgba(255,255,255,0.15);border:2px solid transparent;">DOCX Generator</a>
            <a href="meta-extractor.php" style="padding:8px 20px;border-radius:20px;text-decoration:none;font-size:14px;font-weight:600;color:rgba(255,255,255,0.8);background:rgba(255,255,255,0.15);border:2px solid transparent;">Meta Extractor</a>
            <a href="meta-tool/index.php" style="padding:8px 20px;border-radius:20px;text-decoration:none;font-size:14px;font-weight:600;color:rgba(255,255,255,0.8);background:rgba(255,255,255,0.15);border:2px solid transparent;">Meta Translator</a>
            <a href="lang-generator.php" style="padding:8px 20px;border-radius:20px;text-decoration:none;font-size:14px;font-weight:600;background:white;color:#667eea;border:2px solid white;">Language Tab Generator</a>
        </div>
    </div>

    <div class="main-container">
        <!-- Left Sidebar - Output Directory -->
        <div class="sidebar">
            <div class="sidebar-header">
                <span class="title">Language Documents</span>
                <?php
                $outputDir = __DIR__ . '/output';
                $logFiles = [];
                if (is_dir($outputDir)) {
                    $iterator = new RecursiveIteratorIterator(
                        new RecursiveDirectoryIterator($outputDir, RecursiveDirectoryIterator::SKIP_DOTS),
                        RecursiveIteratorIterator::SELF_FIRST
                    );
                    foreach ($iterator as $file) {
                        if ($file->isFile() && pathinfo($file->getFilename(), PATHINFO_EXTENSION) === 'log') {
                            $relativePath = str_replace($outputDir . '/', '', $file->getPathname());
                            $logFiles[] = 'output/' . $relativePath;
                        }
                    }
                }
                if (!empty($logFiles)) {
                    rsort($logFiles);
                    $latestLog = $logFiles[0];
                    echo '<a href="' . htmlspecialchars($latestLog) . '" target="_blank" class="logs-link">📋 View Latest Log</a>';
                }
                ?>
            </div>
            <div class="filter-bar">
                <input type="text" id="filterInput" placeholder="Filter projects and files..." oninput="applyFilter()">
                <label class="show-hidden-label">
                    <input type="checkbox" id="showHidden" onchange="applyFilter()"> Show hidden
                </label>
            </div>
            <div class="directory-tree" id="directoryTree">
                <?php
                function scanDirectory(string $dir, string $baseDir, array $hiddenProjects = []): void {
                    if (!is_dir($dir)) {
                        echo '<div class="empty-directory">No language documents generated yet</div>';
                        return;
                    }
                    $items = scandir($dir);
                    $hasContent = false;
                    foreach ($items as $item) {
                        if ($item === '.' || $item === '..') continue;
                        $fullPath = $dir . '/' . $item;
                        $relativePath = str_replace($baseDir . '/', '', $fullPath);
                        if (is_dir($fullPath)) {
                            // Only show folder if it contains lang-combined-*.docx files
                            $hasLangFiles = countLangFiles($fullPath);
                            if (!$hasLangFiles) continue;
                            $hasContent = true;
                            $isHiddenAttr = in_array($item, $hiddenProjects) ? 'true' : 'false';
                            echo '<div class="directory-item folder" data-project="' . htmlspecialchars($item) . '" data-hidden="' . $isHiddenAttr . '">';
                            echo '<input type="checkbox" class="folder-checkbox" onclick="event.stopPropagation(); toggleFolderFiles(this);" style="margin-right: 8px;">';
                            echo '<span class="folder-name" onclick="toggleFolder(this.closest(\'.directory-item.folder\'))">' . htmlspecialchars($item) . '</span>';
                            echo '<button class="file-action-btn btn-hide folder-hide-btn" onclick="event.stopPropagation(); toggleHideProject(this, \'' . htmlspecialchars(addslashes($item)) . '\')">Hide</button>';
                            echo '</div>';
                            echo '<div class="folder-content" style="padding-left: 20px;">';
                            scanDirectory($fullPath, $baseDir, $hiddenProjects);
                            echo '</div>';
                        } elseif (pathinfo($item, PATHINFO_EXTENSION) === 'docx' && strpos($item, 'lang-') === 0) {
                            // Only show lang-combined-*.docx files
                            $hasContent = true;
                            echo '<div class="directory-item file">';
                            echo '<div class="file-info">';
                            echo '<input type="checkbox" class="file-checkbox" data-file="' . htmlspecialchars('output/' . $relativePath) . '">';
                            echo '<span class="filename">' . htmlspecialchars(preg_replace('/^lang-/', '', $item)) . '</span>';
                            echo '</div>';
                            echo '<div class="file-actions">';
                            echo '<a href="download.php?file=' . urlencode('output/' . $relativePath) . '" class="file-action-btn btn-download">Download</a>';
                            echo '<a href="remove.php?file=' . urlencode('output/' . $relativePath) . '&redirect=lang-generator.php" onclick="return confirm(\'Delete this file?\')" class="file-action-btn btn-remove">Remove</a>';
                            echo '</div>';
                            echo '</div>';
                        } elseif ($item === 'lang-debug.log') {
                            $hasContent = true;
                            echo '<div class="directory-item file log">';
                            echo '<div class="file-info"><span class="filename">' . htmlspecialchars($item) . '</span></div>';
                            echo '<div class="file-actions">';
                            echo '<a href="output/' . htmlspecialchars($relativePath) . '" target="_blank" class="file-action-btn btn-view">View</a>';
                            echo '</div></div>';
                        }
                    }
                    if (!$hasContent) {
                        echo '<div class="empty-directory">No language documents yet</div>';
                    }
                }

                function countLangFiles(string $dir): bool {
                    foreach (scandir($dir) as $item) {
                        if ($item === '.' || $item === '..') continue;
                        $path = $dir . '/' . $item;
                        if (is_file($path) && pathinfo($item, PATHINFO_EXTENSION) === 'docx' && strpos($item, 'lang-') === 0) {
                            return true;
                        }
                        if (is_dir($path) && countLangFiles($path)) return true;
                    }
                    return false;
                }
                $outputDir = __DIR__ . '/output';
                scanDirectory($outputDir, $outputDir, $hiddenProjects);
                ?>
            </div>
            <div class="bulk-actions">
                <button type="button" class="bulk-action-btn btn-select-all" onclick="toggleSelectAll()">Select All</button>
                <button type="button" class="bulk-action-btn btn-bulk-download" onclick="bulkDownload()" disabled id="bulkDownloadBtn">Download Selected</button>
                <button type="button" class="bulk-action-btn btn-bulk-delete" onclick="bulkDelete()" disabled id="bulkDeleteBtn">Delete Selected</button>
            </div>
        </div>

        <!-- Right Content Area -->
        <div class="content-area">
            <form method="POST" action="lang-processor.php" id="langForm">
                <!-- Project -->
                <?php
                $outputDir = __DIR__ . '/output';
                $existingProjects = [];
                if (is_dir($outputDir)) {
                    foreach (scandir($outputDir) as $item) {
                        if ($item === '.' || $item === '..') continue;
                        $projPath = $outputDir . '/' . $item;
                        if (!is_dir($projPath)) continue;
                        foreach (scandir($projPath) as $file) {
                            if ($file === '.' || $file === '..') continue;
                            if (pathinfo($file, PATHINFO_EXTENSION) === 'docx' && strpos($file, 'lang-') === 0) {
                                $existingProjects[] = $item;
                                break;
                            }
                        }
                    }
                }
                $defaultMode = !empty($existingProjects) ? 'existing' : 'new';
                ?>
                <div class="form-group">
                    <label>Project (Optional)</label>
                    <div class="project-mode-toggle">
                        <label class="project-mode-option">
                            <input type="radio" name="project_mode" value="new" onchange="switchProjectMode('new')"
                                <?php echo $defaultMode === 'new' ? 'checked' : ''; ?>> New project
                        </label>
                        <label class="project-mode-option">
                            <input type="radio" name="project_mode" value="existing" onchange="switchProjectMode('existing')"
                                <?php echo $defaultMode === 'existing' ? 'checked' : ''; ?>> Existing project
                        </label>
                    </div>
                    <div id="projectNewInput" style="<?php echo $defaultMode === 'new' ? '' : 'display:none;'; ?>">
                        <input type="text" id="projectName" name="<?php echo $defaultMode === 'new' ? 'project' : ''; ?>" placeholder="skycop-lang" />
                    </div>
                    <div id="projectExistingInput" style="<?php echo $defaultMode === 'existing' ? '' : 'display:none;'; ?>">
                        <select name="<?php echo $defaultMode === 'existing' ? 'project' : ''; ?>" id="projectSelect">
                            <option value="">— None (root output) —</option>
                            <?php foreach ($existingProjects as $proj): ?>
                            <option value="<?php echo htmlspecialchars($proj); ?>"><?php echo htmlspecialchars($proj); ?></option>
                            <?php endforeach; ?>
                        </select>
                        <?php if (empty($existingProjects)): ?>
                        <div class="help-text" style="margin-top:6px;">No language projects yet — generate your first document to create one.</div>
                        <?php endif; ?>
                    </div>
                    <div class="help-text">Organize generated files in a named subfolder.</div>
                </div>

                <!-- URLs -->
                <div class="form-group urls-group">
                    <label for="urls">Website URLs</label>
                    <div class="textarea-wrapper">
                        <textarea
                            name="urls"
                            id="urls"
                            placeholder="https://example.com/en/page1&#10;https://example.com/es/page1&#10;https://example.com/fr/page1&#10;https://example.com/de/page1&#10;...&#10;Language codes are detected automatically from URL paths."
                            required
                        ></textarea>
                    </div>
                    <div class="help-text">Enter one URL per line. Language codes are detected from the URL path (e.g. /en/, /es/, /fr/). All languages are combined into one output document.</div>
                </div>

                <!-- Language Tabs (auto-populated by JS) -->
                <div class="lang-tabs-section" id="langTabsSection" style="display:none;">
                    <label>Detected Languages</label>
                    <div class="lang-summary" id="langSummary"></div>
                    <div class="lang-tabs-nav" id="langTabsNav"></div>
                    <div id="langTabsContents"></div>
                </div>

                <!-- Selected Selector -->
                <div class="form-group">
                    <label for="selector">Selected Selector (Optional)</label>
                    <input type="text" name="selector" id="selector" placeholder="&lt;article&gt; or .my-class or #my-id" />
                    <div class="help-text">Any tag (&lt;article&gt;, &lt;main&gt;, or just "article"), a class with a dot (.my-class) or an ID with a hash (#my-id). Leave empty for full body.</div>
                </div>

                <!-- Skip Selectors -->
                <div class="form-group">
                    <label for="skip_selectors">Skip Selectors (Optional)</label>
                    <input type="text" name="skip_selectors" id="skip_selectors" placeholder="&lt;header&gt;, .my-class, #my-id" />
                    <div class="help-text">Comma-separated list to exclude — any tag (&lt;header&gt;, &lt;nav&gt;), a class (.my-class) or an ID (#my-id).</div>
                </div>

                <!-- ── Output Destination ─────────────────────────────────── -->
                <div class="form-group output-dest-section">
                    <label>Save Output To</label>

                    <div class="dest-toggle">
                        <!-- DOCX option (always available) -->
                        <div class="dest-option">
                            <input type="radio" name="output_mode" id="modeDOCX" value="docx" checked
                                   onchange="onOutputModeChange()">
                            <label for="modeDOCX">
                                <span class="dest-icon">📄</span> Save as DOCX
                            </label>
                        </div>

                        <!-- Google Sheets option — disabled only if credentials file is missing -->
                        <div class="dest-option">
                            <input type="radio" name="output_mode" id="modeGDrive" value="google_drive"
                                   <?php echo !$googleHasCredentials ? 'disabled' : ''; ?>
                                   onchange="onOutputModeChange()">
                            <label for="modeGDrive" style="<?php echo !$googleHasCredentials ? 'opacity:0.5;cursor:not-allowed;' : ''; ?>">
                                <span class="dest-icon">📄</span> Save to Google Docs
                            </label>
                        </div>
                    </div>

                    <!-- Google Drive details panel (shown when Google Drive is selected) -->
                    <div class="gdrive-panel" id="gdrivePanel">

                        <?php if (!$googleHasCredentials): ?>
                            <div class="google-docs-notice">
                                To enable Google Drive, place <strong>google-credentials.json</strong> in
                                the app root. <a href="https://console.cloud.google.com/" target="_blank"
                                class="google-auth-link">Open Google Cloud Console →</a>
                            </div>

                        <?php elseif (!$googleAuthenticated): ?>
                            <!-- Credentials exist but not yet connected — show Connect button -->
                            <div style="display:flex;align-items:center;gap:10px;margin-bottom:12px;">
                                <div class="google-auth-badge disconnected">
                                    <span class="badge-dot"></span> Not connected
                                </div>
                                <a href="google-auth.php" class="google-auth-link"
                                   style="padding:6px 14px;background:#1a73e8;color:#fff;border-radius:5px;font-size:12px;font-weight:600;text-decoration:none;">
                                    Connect Google Account →
                                </a>
                            </div>
                            <div class="google-docs-notice">
                                Click <strong>Connect Google Account</strong> to authorise access.
                                You'll be redirected back here automatically.
                            </div>

                        <?php else: ?>
                            <div style="display:flex;align-items:center;gap:10px;margin-bottom:12px;">
                                <div class="google-auth-badge connected">
                                    <span class="badge-dot"></span> Connected
                                </div>
                                <a href="google-auth.php?action=revoke" class="google-auth-link"
                                   onclick="return confirm('Disconnect Google account?')">Disconnect</a>
                            </div>

                            <label for="google_doc_title">Document Title (Optional)</label>
                            <input type="text" name="google_doc_title" id="google_doc_title"
                                   placeholder="Leave empty to use the auto-generated filename" />

                            <!-- Destination folder (read from .env — not editable here) -->
                            <div class="gdrive-dest-info">
                                <span class="dest-folder-icon">📁</span>
                                <?php if ($googleDriveFolderId !== ''): ?>
                                    Saves to:&nbsp;
                                    <a href="https://drive.google.com/drive/folders/<?php echo htmlspecialchars($googleDriveFolderId); ?>"
                                       target="_blank">Open destination folder →</a>
                                    &nbsp;<span style="color:#aaa;font-size:11px;">(ID: <?php echo htmlspecialchars($googleDriveFolderId); ?>)</span>
                                <?php else: ?>
                                    Saves to: <strong>My Drive</strong> (root)
                                    &nbsp;<span style="color:#aaa;font-size:11px;">Set <code>GOOGLE_DRIVE_FOLDER_ID</code> in <code>.env</code> to change</span>
                                <?php endif; ?>
                            </div>

                            <div class="help-text" style="margin-top:8px;">
                                One Google Document is created with one tab per detected language.
                                A direct link appears when generation completes.
                            </div>
                        <?php endif; ?>

                    </div><!-- /gdrivePanel -->
                </div>

                <div class="button-group">
                    <button type="submit" id="submitBtn">Generate Combined DOCX</button>
                </div>
            </form>

            <?php
            if (isset($_SESSION['lang_status'])) {
                $status = $_SESSION['lang_status'];
                echo '<div id="statusData" style="display:none;"
                      data-type="' . htmlspecialchars($status['type']) . '"
                      data-message="' . htmlspecialchars($status['message']) . '"';
                if (isset($status['processed'], $status['total'])) {
                    echo ' data-processed="' . $status['processed'] . '" data-total="' . $status['total'] . '"';
                }
                if (isset($status['log_file'])) {
                    echo ' data-log-file="' . htmlspecialchars($status['log_file']) . '"';
                }
                if (isset($status['google_docs'])) {
                    echo ' data-google-docs="' . htmlspecialchars(json_encode($status['google_docs'])) . '"';
                }
                if (isset($status['google_folder_url'])) {
                    echo ' data-google-folder-url="' . htmlspecialchars($status['google_folder_url']) . '"';
                }
                if (isset($status['google_docs_error'])) {
                    echo ' data-google-docs-error="' . htmlspecialchars($status['google_docs_error']) . '"';
                }
                echo '></div>';
                unset($_SESSION['lang_status']);
            }
            ?>
        </div>
    </div>

    <script>
        // ISO 639-1 language codes (all 184 codes)
        const ISO_LANG_CODES = [
            'ab','aa','af','ak','sq','am','ar','an','hy','as','av','ae','ay','az',
            'bm','ba','eu','be','bn','bh','bi','bs','br','bg','my','ca','ch','ce',
            'ny','zh','cv','kw','co','cr','hr','cs','da','dv','nl','dz','en','eo',
            'et','ee','fo','fj','fi','fr','ff','gl','ka','de','el','gn','gu','ht',
            'ha','he','hz','hi','ho','hu','ia','id','ie','ga','ig','ik','io','is',
            'it','iu','ja','jv','kl','kn','kr','ks','kk','km','ki','rw','ky','kv',
            'kg','ko','ku','kj','la','lb','lg','li','ln','lo','lt','lu','lv','gv',
            'mk','mg','ms','ml','mt','mi','mr','mh','mn','na','nv','nd','ne','ng',
            'nb','nn','no','ii','nr','oc','oj','cu','om','or','os','pa','pi','fa',
            'pl','ps','pt','qu','rm','rn','ro','ru','sa','sc','sd','se','sm','sg',
            'sr','gd','sn','si','sk','sl','so','st','es','su','sw','ss','sv','ta',
            'te','tg','th','ti','bo','tk','tl','tn','to','tr','ts','tt','tw','ty',
            'ug','uk','ur','uz','ve','vi','vo','wa','cy','wo','fy','xh','yi','yo',
            'za','zu'
        ];

        // Language display names for known codes
        const LANG_NAMES = {
            'en':'English','es':'Spanish','fr':'French','de':'German','it':'Italian',
            'pt':'Portuguese','nl':'Dutch','pl':'Polish','ru':'Russian','ar':'Arabic',
            'zh':'Chinese','ja':'Japanese','ko':'Korean','sv':'Swedish','da':'Danish',
            'fi':'Finnish','nb':'Norwegian','no':'Norwegian','cs':'Czech','sk':'Slovak',
            'ro':'Romanian','hu':'Hungarian','bg':'Bulgarian','hr':'Croatian','sr':'Serbian',
            'uk':'Ukrainian','el':'Greek','tr':'Turkish','he':'Hebrew','fa':'Persian',
            'hi':'Hindi','bn':'Bengali','th':'Thai','vi':'Vietnamese','id':'Indonesian',
            'ms':'Malay','ca':'Catalan','eu':'Basque','gl':'Galician','af':'Afrikaans',
            'sq':'Albanian','hy':'Armenian','ka':'Georgian','lv':'Latvian','lt':'Lithuanian',
            'et':'Estonian','sl':'Slovenian','mk':'Macedonian','is':'Icelandic',
            'ga':'Irish','cy':'Welsh','mt':'Maltese','lb':'Luxembourgish'
        };

        function detectUrlLang(url) {
            try {
                const u = new URL(url.trim());
                const parts = u.pathname.split('/').filter(p => p.length > 0);
                if (parts.length) {
                    const first = parts[0].toLowerCase();

                    // Any 2-letter code (language OR region, e.g. gb, fr, de, us)
                    if (/^[a-z]{2}$/.test(first)) {
                        return first;
                    }

                    // Locale code: en-us, pt-br, zh-cn, en-gb, etc.
                    if (/^[a-z]{2}[-_][a-z]{2,4}$/i.test(first)) {
                        return first.toLowerCase().replace('_', '-');
                    }
                }
            } catch(e) {}
            return 'en';
        }

        function getLangLabel(code) {
            if (code === 'default') return 'Default';
            const base = code.split('-')[0];
            const name = LANG_NAMES[base] || LANG_NAMES[code];
            if (name) return code.toUpperCase() + ' · ' + name;
            return code.toUpperCase(); // e.g. GB, US — show as-is
        }

        let activeTab = null;

        function updateLangTabs() {
            const text = document.getElementById('urls').value;
            const lines = text.split('\n').map(l => l.trim()).filter(l => l.length > 0);

            const groups = {};
            lines.forEach(url => {
                if (!url.startsWith('http')) return;
                const key = detectUrlLang(url); // always returns a code ('en' as default)
                if (!groups[key]) groups[key] = [];
                groups[key].push(url);
            });

            const section = document.getElementById('langTabsSection');
            const nav = document.getElementById('langTabsNav');
            const contents = document.getElementById('langTabsContents');
            const summary = document.getElementById('langSummary');

            const keys = Object.keys(groups);
            if (keys.length === 0) {
                section.style.display = 'none';
                return;
            }

            section.style.display = '';

            // Sort language codes alphabetically
            const sortedKeys = keys.sort((a, b) => a.localeCompare(b));

            const langCount = sortedKeys.filter(k => k !== '__other__').length;
            const totalUrls = lines.filter(l => l.startsWith('http')).length;
            summary.innerHTML = `Detected <strong>${langCount}</strong> language${langCount !== 1 ? 's' : ''} across <strong>${totalUrls}</strong> URL${totalUrls !== 1 ? 's' : ''} — all will be combined into one document.`;

            nav.innerHTML = '';
            contents.innerHTML = '';

            // Keep active tab if still valid, else use first
            if (!activeTab || !groups[activeTab]) {
                activeTab = sortedKeys[0];
            }

            sortedKeys.forEach(key => {
                const urls = groups[key];
                const label = getLangLabel(key);

                // Tab button
                const btn = document.createElement('button');
                btn.type = 'button';
                btn.className = 'lang-tab-btn' + (key === activeTab ? ' active' : '');
                btn.dataset.lang = key;
                btn.innerHTML = label + ' <span class="lang-tab-badge">' + urls.length + '</span>';
                btn.addEventListener('click', () => switchLangTab(key));
                nav.appendChild(btn);

                // Tab content panel
                const panel = document.createElement('div');
                panel.className = 'lang-tab-content' + (key === activeTab ? ' active' : '');
                panel.id = 'lc-' + key;

                const urlList = document.createElement('div');
                urlList.className = 'lang-urls-panel';
                if (urls.length === 0) {
                    urlList.innerHTML = '<div class="lang-empty">No URLs detected for this language.</div>';
                } else {
                    urls.forEach(u => {
                        const item = document.createElement('div');
                        item.className = 'lang-url-item';
                        item.title = u;
                        item.textContent = u;
                        urlList.appendChild(item);
                    });
                }
                panel.appendChild(urlList);
                contents.appendChild(panel);
            });
        }

        function switchLangTab(key) {
            activeTab = key;
            document.querySelectorAll('.lang-tab-btn').forEach(btn => {
                btn.classList.toggle('active', btn.dataset.lang === key);
            });
            document.querySelectorAll('.lang-tab-content').forEach(panel => {
                panel.classList.toggle('active', panel.id === 'lc-' + key);
            });
        }

        document.getElementById('urls').addEventListener('input', updateLangTabs);

        // Project mode toggle
        function switchProjectMode(mode) {
            const newInput = document.getElementById('projectNewInput');
            const existingInput = document.getElementById('projectExistingInput');
            const nameField = document.getElementById('projectName');
            const selectField = document.getElementById('projectSelect');
            if (mode === 'existing') {
                newInput.style.display = 'none';
                existingInput.style.display = '';
                if (nameField) nameField.removeAttribute('name');
                if (selectField) selectField.setAttribute('name', 'project');
            } else {
                newInput.style.display = '';
                existingInput.style.display = 'none';
                if (nameField) nameField.setAttribute('name', 'project');
                if (selectField) selectField.removeAttribute('name');
            }
        }

        // Sidebar: hidden projects state
        let hiddenProjects = <?php echo json_encode($hiddenProjects); ?>;

        function toggleHideProject(button, projectName) {
            const folderItem = button.closest('.directory-item.folder');
            const isCurrentlyHidden = folderItem.dataset.hidden === 'true';
            if (isCurrentlyHidden) {
                hiddenProjects = hiddenProjects.filter(p => p !== projectName);
                folderItem.dataset.hidden = 'false';
            } else {
                if (!hiddenProjects.includes(projectName)) hiddenProjects.push(projectName);
                folderItem.dataset.hidden = 'true';
            }
            fetch('settings.php', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ action: 'save_settings', hidden_projects: hiddenProjects })
            }).catch(() => {});
            applyFilter();
        }

        function applyFilter() {
            const filter = document.getElementById('filterInput').value.toLowerCase().trim();
            const showHidden = document.getElementById('showHidden').checked;
            document.querySelectorAll('#directoryTree .directory-item.folder').forEach(folder => {
                const folderContent = folder.nextElementSibling;
                const hasFolderContent = folderContent && folderContent.classList.contains('folder-content');
                const isHidden = folder.dataset.hidden === 'true';
                const projectName = (folder.dataset.project || '').toLowerCase();
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
                    const folderMatches = projectName.includes(filter);
                    let hasMatchingFiles = false;
                    if (hasFolderContent) {
                        folderContent.querySelectorAll('.directory-item.file').forEach(fileItem => {
                            const fileName = (fileItem.querySelector('.filename')?.textContent || '').toLowerCase();
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

        function toggleFolder(element) {
            const folderContent = element.nextElementSibling;
            if (folderContent && folderContent.classList.contains('folder-content')) {
                folderContent.style.display = folderContent.style.display === 'none' ? 'block' : 'none';
            }
        }

        function toggleFolderFiles(folderCheckbox) {
            const folderItem = folderCheckbox.closest('.directory-item.folder');
            const folderContent = folderItem.nextElementSibling;
            if (folderContent && folderContent.classList.contains('folder-content')) {
                folderContent.querySelectorAll('.file-checkbox').forEach(cb => {
                    cb.checked = folderCheckbox.checked;
                });
                updateBulkActionButtons();
            }
        }

        function updateBulkActionButtons() {
            const checkedBoxes = document.querySelectorAll('.file-checkbox:checked');
            const downloadBtn = document.getElementById('bulkDownloadBtn');
            const deleteBtn = document.getElementById('bulkDeleteBtn');
            downloadBtn.disabled = checkedBoxes.length === 0;
            deleteBtn.disabled = checkedBoxes.length === 0;
        }

        function toggleSelectAll() {
            const checkboxes = document.querySelectorAll('.file-checkbox');
            const allChecked = Array.from(checkboxes).every(cb => cb.checked);
            checkboxes.forEach(cb => { cb.checked = !allChecked; });
            updateBulkActionButtons();
        }

        function bulkDownload() {
            const checkedBoxes = document.querySelectorAll('.file-checkbox:checked');
            if (!checkedBoxes.length) return;
            checkedBoxes.forEach((checkbox, index) => {
                setTimeout(() => {
                    const iframe = document.createElement('iframe');
                    iframe.style.display = 'none';
                    iframe.src = 'download.php?file=' + encodeURIComponent(checkbox.dataset.file);
                    document.body.appendChild(iframe);
                    setTimeout(() => document.body.removeChild(iframe), 1000);
                }, index * 500);
            });
            showToast('success', `Downloading ${checkedBoxes.length} file(s)...`);
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
            .catch(() => showToast('error', 'An error occurred'));
        }

        function deleteFolder(folderPath) {
            if (!confirm(`Delete folder "${folderPath}" and all its contents?`)) return;
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
            .catch(() => showToast('error', 'An error occurred'));
        }

        document.addEventListener('DOMContentLoaded', function() {
            document.querySelectorAll('.file-checkbox').forEach(cb => {
                cb.addEventListener('change', updateBulkActionButtons);
            });
            applyFilter();

            // Sync project mode display with whichever radio the browser has checked
            // (browser may restore a cached selection that differs from server default)
            const checkedRadio = document.querySelector('input[name="project_mode"]:checked');
            if (checkedRadio) switchProjectMode(checkedRadio.value);

            // Restore status toast from session
            const statusData = document.getElementById('statusData');
            if (statusData) {
                const type = statusData.dataset.type;
                const message = statusData.dataset.message;
                const options = {};
                if (statusData.dataset.processed && statusData.dataset.total) {
                    options.processed = parseInt(statusData.dataset.processed);
                    options.total = parseInt(statusData.dataset.total);
                }
                if (statusData.dataset.logFile) options.logFile = statusData.dataset.logFile;
                if (statusData.dataset.googleDocUrl) options.googleDocUrl = statusData.dataset.googleDocUrl;
                if (statusData.dataset.googleFolderUrl) options.googleFolderUrl = statusData.dataset.googleFolderUrl;
                showToast(type, message, options);

                // Show a separate Google Docs error toast if export failed
                if (statusData.dataset.googleDocsError) {
                    showToast('error', 'Google Docs export failed: ' + statusData.dataset.googleDocsError);
                }

                statusData.remove();
            }
        });

        // Toast system
        function showToast(type, message, options = {}) {
            const container = document.getElementById('toastContainer');
            const toast = document.createElement('div');
            toast.className = `toast toast-${type}`;
            const icons = { success: '✓', error: '✕', processing: 'ℹ' };
            const titles = { success: 'Success', error: 'Error', processing: 'Processing' };
            let html = `<span class="toast-icon">${icons[type] || 'ℹ'}</span>
                <div class="toast-content">
                    <div class="toast-title">${titles[type] || 'Notification'}</div>
                    <div class="toast-message">${message}</div>`;
            if (options.processed !== undefined && options.total !== undefined) {
                const pct = (options.processed / options.total) * 100;
                html += `<div class="toast-progress"><div class="toast-progress-bar" style="width:${pct}%"></div></div>`;
            }
            if (options.logFile) {
                html += `<div style="margin-top:8px;"><a href="${options.logFile}" target="_blank" style="color:#667eea;text-decoration:underline;font-weight:500;">View Error Log</a></div>`;
            }
            if (options.googleDocUrl) {
                html += `<div style="margin-top:8px;display:flex;flex-direction:column;gap:4px;">`;
                html += `<a href="${options.googleDocUrl}" target="_blank" style="color:#1a73e8;text-decoration:underline;font-weight:600;">📄 Open Google Document →</a>`;
                if (options.googleFolderUrl) {
                    html += `<a href="${options.googleFolderUrl}" target="_blank" style="color:#555;text-decoration:underline;font-size:12px;">📁 Open folder in Drive →</a>`;
                }
                html += `</div>`;
            }
            html += `</div><button class="toast-close" onclick="closeToast(this)">×</button>`;
            toast.innerHTML = html;
            container.appendChild(toast);
            const duration = options.duration || (type === 'error' ? 8000 : 5000);
            setTimeout(() => closeToast(toast.querySelector('.toast-close')), duration);
        }

        function closeToast(button) {
            const toast = button.closest('.toast');
            toast.classList.add('hiding');
            setTimeout(() => toast.remove(), 300);
        }

        function onOutputModeChange() {
            const mode    = document.querySelector('input[name="output_mode"]:checked')?.value || 'docx';
            const panel   = document.getElementById('gdrivePanel');
            const btn     = document.getElementById('submitBtn');

            if (mode === 'google_drive') {
                panel.classList.add('active');
                btn.textContent = 'Generate & Save to Google Docs';
            } else {
                panel.classList.remove('active');
                btn.textContent = 'Generate Combined DOCX';
            }
        }

        const GDRIVE_AUTHENTICATED = <?php echo $googleAuthenticated ? 'true' : 'false'; ?>;

        document.getElementById('langForm').addEventListener('submit', function(e) {
            const mode = document.querySelector('input[name="output_mode"]:checked')?.value || 'docx';

            // Block submission if Google Drive selected but not authenticated
            if (mode === 'google_drive' && !GDRIVE_AUTHENTICATED) {
                e.preventDefault();
                showToast('error', 'Please connect your Google account first before saving to Google Drive.');
                return;
            }

            document.getElementById('submitBtn').disabled = true;
            if (mode === 'google_drive') {
                document.getElementById('submitBtn').textContent = 'Saving to Google Docs…';
                showToast('processing', 'Creating Google Document with language tabs…', { duration: 10000 });
            } else {
                document.getElementById('submitBtn').textContent = 'Processing…';
                showToast('processing', 'Generating combined language document…', { duration: 3000 });
            }
        });

        // Show Google auth status toasts
        (function () {
            const params = new URLSearchParams(window.location.search);
            const status = params.get('google_status');
            if (!status) return;

            const messages = {
                connected:       { type: 'success', msg: 'Google account connected successfully.' },
                revoked:         { type: 'success', msg: 'Google account disconnected.' },
                no_credentials:  { type: 'error',   msg: 'Google credentials file not found. Please add google-credentials.json.' },
                auth_failed:     { type: 'error',   msg: 'Google authentication failed: ' + (params.get('error') || 'unknown error') },
            };

            const entry = messages[status];
            if (entry) showToast(entry.type, entry.msg);

            // Clean the URL so the toast doesn't reappear on reload
            const clean = window.location.pathname;
            history.replaceState(null, '', clean);
        })();
    </script>
</body>
</html>
