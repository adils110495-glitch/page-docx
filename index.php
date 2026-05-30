<?php
session_start();

// Load hidden project settings
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
    <title>Website to DOCX Generator</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
        }

        .header {
            text-align: center;
            color: white;
            margin-bottom: 20px;
        }

        .header h1 {
            font-size: 32px;
            margin-bottom: 8px;
        }

        .header p {
            font-size: 16px;
            opacity: 0.9;
        }

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
            box-shadow: 0 20px 60px rgba(0, 0, 0, 0.3);
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

        .sidebar-header .title {
            flex: 1;
        }

        .logs-link {
            color: white;
            text-decoration: none;
            font-size: 14px;
            font-weight: 500;
            padding: 6px 12px;
            background: rgba(255, 255, 255, 0.2);
            border-radius: 4px;
            transition: background 0.2s;
        }

        .logs-link:hover {
            background: rgba(255, 255, 255, 0.3);
        }

        .directory-tree {
            flex: 1;
            overflow-y: auto;
            padding: 15px;
        }

        .directory-item {
            padding: 10px;
            margin-bottom: 5px;
            border-radius: 6px;
            cursor: pointer;
            transition: background 0.2s;
            font-size: 14px;
        }

        .directory-item:hover {
            background: #f0f0f0;
        }

        .directory-item.folder {
            font-weight: 600;
            color: #667eea;
        }

        .directory-item.file {
            padding-left: 30px;
            color: #666;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }

        .directory-item.file .filename {
            flex: 1;
        }

        .directory-item.file .filename::before {
            content: "📄 ";
        }

        .directory-item.log .filename::before {
            content: "📋 ";
        }

        .directory-item.folder::before {
            content: "📁 ";
        }

        .directory-item.folder {
            display: flex;
            align-items: center;
        }

        .directory-item.folder .folder-name {
            flex: 1;
            cursor: pointer;
        }

        .directory-item.folder .folder-delete-btn {
            display: none;
            margin-left: 10px;
        }

        .directory-item.folder:hover .folder-delete-btn {
            display: inline-block;
        }

        .file-actions {
            display: none;
            gap: 8px;
        }

        .directory-item.file:hover .file-actions {
            display: flex;
        }

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

        .btn-download {
            background: #28a745;
        }

        .btn-view {
            background: #007bff;
        }

        .btn-remove {
            background: #dc3545;
        }

        .empty-directory {
            text-align: center;
            padding: 40px 20px;
            color: #999;
            font-style: italic;
        }

        .bulk-actions {
            padding: 15px;
            border-top: 1px solid rgba(255, 255, 255, 0.2);
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

        .bulk-action-btn:hover {
            opacity: 0.9;
        }

        .bulk-action-btn:disabled {
            opacity: 0.5;
            cursor: not-allowed;
        }

        .btn-bulk-download {
            background: #28a745;
            color: white;
        }

        .btn-bulk-delete {
            background: #dc3545;
            color: white;
        }

        .btn-select-all {
            background: rgba(255, 255, 255, 0.2);
            color: white;
            border: 1px solid rgba(255, 255, 255, 0.3);
        }

        .file-checkbox {
            margin-right: 8px;
            cursor: pointer;
            width: 16px;
            height: 16px;
        }

        .directory-item.file {
            padding-left: 10px;
            color: #666;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }

        .file-info {
            display: flex;
            align-items: center;
            flex: 1;
        }

        .content-area {
            background: white;
            border-radius: 12px;
            box-shadow: 0 20px 60px rgba(0, 0, 0, 0.3);
            padding: 30px 40px;
            display: flex;
            flex-direction: column;
        }

        .form-wrapper {
            flex: 1;
            display: flex;
            flex-direction: column;
            min-height: 0;
        }

        #generatorForm {
            display: flex;
            flex-direction: column;
            height: 100%;
        }

        .form-group {
            margin-bottom: 20px;
            flex-shrink: 0;
        }

        .form-group.urls-group {
            flex: 1;
            display: flex;
            flex-direction: column;
            min-height: 0;
            margin-bottom: 16px;
        }

        .urls-group .textarea-wrapper {
            flex: 1;
            display: flex;
            flex-direction: column;
            min-height: 0;
        }

        label {
            display: block;
            color: #333;
            font-weight: 600;
            margin-bottom: 6px;
            font-size: 13px;
        }

        .help-text {
            font-size: 11px;
            color: #777;
            margin-top: 5px;
            font-style: italic;
            line-height: 1.3;
        }

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
            max-height: none;
            box-sizing: border-box;
            line-height: 1.5;
        }

        textarea::placeholder {
            font-size: 12px;
            line-height: 1.6;
        }

        textarea:focus {
            outline: none;
            border-color: #667eea;
        }

        input[type="text"] {
            width: 100%;
            padding: 10px 12px;
            border: 2px solid #e0e0e0;
            border-radius: 6px;
            font-size: 13px;
            transition: border-color 0.3s;
            box-sizing: border-box;
        }

        input[type="text"]:focus {
            outline: none;
            border-color: #667eea;
        }

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
           /*  width: 100%; */
            flex-shrink: 0;
        }

        button:hover {
            transform: translateY(-2px);
            box-shadow: 0 10px 20px rgba(102, 126, 234, 0.4);
        }

        button:active {
            transform: translateY(0);
        }

        button:disabled {
            opacity: 0.6;
            cursor: not-allowed;
            transform: none;
        }

        .button-group {
            flex-shrink: 0;
        }

        .status-message {
            margin-top: 20px;
            padding: 15px;
            border-radius: 6px;
            font-size: 14px;
        }

        .status-success {
            background: #d4edda;
            color: #155724;
            border: 1px solid #c3e6cb;
        }

        .status-error {
            background: #f8d7da;
            color: #721c24;
            border: 1px solid #f5c6cb;
        }

        .status-processing {
            background: #d1ecf1;
            color: #0c5460;
            border: 1px solid #bee5eb;
        }

        .progress-bar {
            width: 100%;
            height: 30px;
            background: #e0e0e0;
            border-radius: 15px;
            overflow: hidden;
            margin-top: 10px;
        }

        .progress-fill {
            height: 100%;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            transition: width 0.3s;
            display: flex;
            align-items: center;
            justify-content: center;
            color: white;
            font-weight: 600;
            font-size: 12px;
        }

        @media (max-width: 1024px) {
            .main-container {
                grid-template-columns: 1fr;
                height: auto;
            }

            .sidebar {
                max-height: 300px;
            }
        }

        /* Toaster Notification Styles */
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
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
            padding: 16px 20px;
            display: flex;
            align-items: flex-start;
            gap: 12px;
            animation: slideIn 0.3s ease-out;
            position: relative;
            overflow: hidden;
        }

        .toast::before {
            content: '';
            position: absolute;
            left: 0;
            top: 0;
            bottom: 0;
            width: 4px;
        }

        .toast.toast-success::before {
            background: #28a745;
        }

        .toast.toast-error::before {
            background: #dc3545;
        }

        .toast.toast-processing::before {
            background: #17a2b8;
        }

        .toast-icon {
            font-size: 20px;
            flex-shrink: 0;
            line-height: 1;
        }

        .toast-content {
            flex: 1;
            font-size: 14px;
            color: #333;
            line-height: 1.5;
        }

        .toast-title {
            font-weight: 600;
            margin-bottom: 4px;
        }

        .toast-message {
            font-size: 13px;
            color: #666;
        }

        .toast-close {
            background: none;
            border: none;
            color: #999;
            cursor: pointer;
            font-size: 18px;
            padding: 0;
            width: 20px;
            height: 20px;
            flex-shrink: 0;
            line-height: 1;
            transition: color 0.2s;
        }

        .toast-close:hover {
            color: #333;
        }

        .toast-progress {
            margin-top: 8px;
            height: 4px;
            background: #e0e0e0;
            border-radius: 2px;
            overflow: hidden;
        }

        .toast-progress-bar {
            height: 100%;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            transition: width 0.3s;
        }

        @keyframes slideIn {
            from {
                transform: translateX(400px);
                opacity: 0;
            }
            to {
                transform: translateX(0);
                opacity: 1;
            }
        }

        @keyframes slideOut {
            from {
                transform: translateX(0);
                opacity: 1;
            }
            to {
                transform: translateX(400px);
                opacity: 0;
            }
        }

        .toast.hiding {
            animation: slideOut 0.3s ease-out forwards;
        }

        /* Project mode toggle */
        .project-mode-toggle {
            display: flex;
            gap: 16px;
            margin-bottom: 8px;
        }

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

        .project-mode-option input[type="radio"] {
            width: auto;
            margin: 0;
            cursor: pointer;
        }

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

        select:focus {
            outline: none;
            border-color: #667eea;
        }

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

        .filter-bar input[type="text"]:focus {
            outline: none;
            border-color: #667eea;
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

        /* Folder hide button */
        .folder-hide-btn {
            display: none;
            background: #6c757d !important;
            margin-left: 5px;
        }

        .directory-item.folder:hover .folder-hide-btn {
            display: inline-block;
        }

        .directory-item.folder.is-hidden-project {
            opacity: 0.55;
            border-left: 3px solid #aaa;
            padding-left: 7px;
        }

        .directory-item.folder.is-hidden-project .folder-hide-btn {
            display: inline-block;
            background: #28a745 !important;
        }
    </style>
</head>
<body>
    <!-- Toast Notification Container -->
    <div class="toast-container" id="toastContainer"></div>

    <div class="header">
        <h1>Website to DOCX Generator</h1>
        <p>Convert website pages into formatted DOCX documents</p>
        <div style="display:flex;justify-content:center;gap:8px;margin-top:14px;">
            <a href="index.php" style="padding:8px 20px;border-radius:20px;text-decoration:none;font-size:14px;font-weight:600;background:white;color:#667eea;border:2px solid white;">DOCX Generator</a>
            <a href="meta-extractor.php" style="padding:8px 20px;border-radius:20px;text-decoration:none;font-size:14px;font-weight:600;color:rgba(255,255,255,0.8);background:rgba(255,255,255,0.15);border:2px solid transparent;">Meta Extractor</a>
            <a href="meta-tool/index.php" style="padding:8px 20px;border-radius:20px;text-decoration:none;font-size:14px;font-weight:600;color:rgba(255,255,255,0.8);background:rgba(255,255,255,0.15);border:2px solid transparent;">Meta Translator</a>
            <a href="lang-generator.php" style="padding:8px 20px;border-radius:20px;text-decoration:none;font-size:14px;font-weight:600;color:rgba(255,255,255,0.8);background:rgba(255,255,255,0.15);border:2px solid transparent;">Language Tab Generator</a>
        </div>
    </div>

    <div class="main-container">
        <!-- Left Sidebar - Directory Structure -->
        <div class="sidebar">
            <div class="sidebar-header">
                <span class="title">Output Directory</span>
                <?php
                // Find all log files
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
                    // Get the most recent log file
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
                // Returns true if the folder has at least one regular (non-lang) DOCX or log file
                function hasDocxFiles($dir) {
                    if (!is_dir($dir)) return false;
                    foreach (scandir($dir) as $file) {
                        if ($file === '.' || $file === '..') continue;
                        $path = $dir . '/' . $file;
                        if (is_dir($path) && hasDocxFiles($path)) return true;
                        if (!is_file($path)) continue;
                        $ext = pathinfo($file, PATHINFO_EXTENSION);
                        if ($ext === 'docx' && strpos($file, 'lang-') !== 0) return true;
                        if ($ext === 'log' && $file !== 'lang-debug.log') return true;
                    }
                    return false;
                }

                function scanDirectory($dir, $baseDir, $hiddenProjects = []) {
                    if (!is_dir($dir)) {
                        echo '<div class="empty-directory">No files generated yet</div>';
                        return;
                    }

                    $items = scandir($dir);
                    $hasContent = false;

                    foreach ($items as $item) {
                        if ($item === '.' || $item === '..') continue;

                        $fullPath = $dir . '/' . $item;
                        $relativePath = str_replace($baseDir . '/', '', $fullPath);

                        if (is_dir($fullPath)) {
                            // Skip folders that contain only lang-tab-generator files
                            if (!hasDocxFiles($fullPath)) continue;
                            $hasContent = true;
                            $isHiddenAttr = in_array($item, $hiddenProjects) ? 'true' : 'false';
                            echo '<div class="directory-item folder" data-project="' . htmlspecialchars($item) . '" data-hidden="' . $isHiddenAttr . '">';
                            echo '<input type="checkbox" class="folder-checkbox" onclick="event.stopPropagation(); toggleFolderFiles(this);" style="margin-right: 8px;">';
                            echo '<span class="folder-name" onclick="toggleFolder(this.closest(\'.directory-item.folder\'))">' . htmlspecialchars($item) . '</span>';
                            echo '<button class="file-action-btn btn-hide folder-hide-btn" onclick="event.stopPropagation(); toggleHideProject(this, \'' . htmlspecialchars(addslashes($item)) . '\')">Hide</button>';
                            echo '<button class="file-action-btn btn-remove folder-delete-btn" onclick="event.stopPropagation(); deleteFolder(\'' . htmlspecialchars('output/' . $relativePath) . '\')">Delete</button>';
                            echo '</div>';
                            echo '<div class="folder-content" style="padding-left: 20px;">';
                            scanDirectory($fullPath, $baseDir, $hiddenProjects);
                            echo '</div>';
                        } elseif (pathinfo($item, PATHINFO_EXTENSION) === 'docx' && strpos($item, 'lang-') !== 0) {
                            $hasContent = true;
                            echo '<div class="directory-item file">';
                            echo '<div class="file-info">';
                            echo '<input type="checkbox" class="file-checkbox" data-file="' . htmlspecialchars('output/' . $relativePath) . '">';
                            echo '<span class="filename">' . htmlspecialchars($item) . '</span>';
                            echo '</div>';
                            echo '<div class="file-actions">';
                            echo '<a href="download.php?file=' . urlencode('output/' . $relativePath) . '" class="file-action-btn btn-download">Download</a>';
                            echo '<a href="remove.php?file=' . urlencode('output/' . $relativePath) . '" onclick="return confirm(\'Delete this file?\')" class="file-action-btn btn-remove">Remove</a>';
                            echo '</div>';
                            echo '</div>';
                        } elseif (pathinfo($item, PATHINFO_EXTENSION) === 'log' && $item !== 'lang-debug.log') {
                            $hasContent = true;
                            echo '<div class="directory-item file log">';
                            echo '<div class="file-info">';
                            echo '<span class="filename">' . htmlspecialchars($item) . '</span>';
                            echo '</div>';
                            echo '<div class="file-actions">';
                            echo '<a href="output/' . htmlspecialchars($relativePath) . '" target="_blank" class="file-action-btn btn-view">View</a>';
                            echo '<a href="remove.php?file=' . urlencode('output/' . $relativePath) . '" onclick="return confirm(\'Delete this log file?\')" class="file-action-btn btn-remove">Remove</a>';
                            echo '</div>';
                            echo '</div>';
                        }
                    }

                    if (!$hasContent) {
                        echo '<div class="empty-directory">Empty directory</div>';
                    }
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

        <!-- Right Content Area - Form -->
        <div class="content-area">
            <form method="POST" action="generator.php" id="generatorForm">
                <div class="form-group">
                    <label>Project (Optional)</label>
                    <div class="project-mode-toggle">
                        <label class="project-mode-option">
                            <input type="radio" name="project_mode" value="new" onchange="switchProjectMode('new')"> New project
                        </label>
                        <label class="project-mode-option">
                            <input type="radio" name="project_mode" value="existing" checked onchange="switchProjectMode('existing')"> Existing project
                        </label>
                    </div>
                    <div id="projectNewInput" style="display:none;">
                        <input
                            type="text"
                            id="projectName"
                            placeholder="skycop-fr"
                        />
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
                        <div class="help-text" style="padding: 8px 0;">No existing projects found. Use "New project" to create one.</div>
                        <?php endif; ?>
                    </div>
                    <div class="help-text">
                        Organize generated files in a named subfolder. If empty or none selected, files go to the root output directory.
                    </div>
                </div>

                <div class="form-group urls-group">
                    <label for="urls">Website URLs</label>
                    <div class="textarea-wrapper">
                        <textarea
                            name="urls"
                            id="urls"
                            placeholder="https://example.com/page1&#10;https://example.com/page2&#10;https://example.com/page3&#10;...&#10;(Up to 100 URLs)"
                            required
                        ></textarea>
                    </div>
                    <div class="help-text">Enter one URL per line (http:// or https://). Maximum 100 URLs per batch.</div>
                </div>

                <div class="form-group">
                    <label for="selector">CSS Selector (Optional)</label>
                    <input
                        type="text"
                        name="selector"
                        id="selector"
                        placeholder="your_right_contents"
                    />
                    <div class="help-text">
                        Enter a tag name, class name, or ID (e.g., "main", "your_right_contents", "#content").
                        Works with any HTML tag (div, main, section, article, etc.). If empty, full &lt;body&gt; content will be extracted.
                    </div>
                </div>

                <div class="form-group">
                    <label for="skip_selectors">Skip Selectors (Optional)</label>
                    <input
                        type="text"
                        name="skip_selectors"
                        id="skip_selectors"
                        placeholder="header, footer, nav, sidebar, ads"
                    />
                    <div class="help-text">
                        Enter CSS class names or element IDs to exclude from content (comma-separated).
                        Examples: "header, footer, sidebar, nav, ads" or ".menu, #sidebar, .advertisement"
                    </div>
                </div>

                <div class="button-group">
                    <button type="submit" id="submitBtn">Generate DOCX</button>
                </div>
            </form>

            <?php
            if (isset($_SESSION['status'])) {
                $status = $_SESSION['status'];
                // Store status data for JavaScript to display as toast
                echo '<div id="statusData" style="display:none;"
                      data-type="' . htmlspecialchars($status['type']) . '"
                      data-message="' . htmlspecialchars($status['message']) . '"';

                if (isset($status['processed']) && isset($status['total'])) {
                    echo ' data-processed="' . $status['processed'] . '"
                          data-total="' . $status['total'] . '"';
                }

                if (isset($status['log_file'])) {
                    echo ' data-log-file="' . htmlspecialchars($status['log_file']) . '"';
                }

                echo '></div>';

                unset($_SESSION['status']);
            }
            ?>
        </div>
    </div>

    <script>
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

        // Hidden projects state (loaded from PHP)
        let hiddenProjects = <?php echo json_encode($hiddenProjects); ?>;

        function toggleHideProject(button, projectName) {
            const folderItem = button.closest('.directory-item.folder');
            const isCurrentlyHidden = folderItem.dataset.hidden === 'true';

            if (isCurrentlyHidden) {
                hiddenProjects = hiddenProjects.filter(p => p !== projectName);
                folderItem.dataset.hidden = 'false';
            } else {
                if (!hiddenProjects.includes(projectName)) {
                    hiddenProjects.push(projectName);
                }
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

        function applyFilter() {
            const filter = document.getElementById('filterInput').value.toLowerCase().trim();
            const showHidden = document.getElementById('showHidden').checked;

            const folderItems = document.querySelectorAll('#directoryTree .directory-item.folder');

            folderItems.forEach(folder => {
                const folderContent = folder.nextElementSibling;
                const hasFolderContent = folderContent && folderContent.classList.contains('folder-content');
                const isHidden = folder.dataset.hidden === 'true';
                const projectName = (folder.dataset.project || '').toLowerCase();

                // Update hide button text and color
                const hideBtn = folder.querySelector('.folder-hide-btn');
                if (hideBtn) {
                    hideBtn.textContent = isHidden ? 'Unhide' : 'Hide';
                }

                // Update hidden indicator class
                if (isHidden && showHidden) {
                    folder.classList.add('is-hidden-project');
                } else {
                    folder.classList.remove('is-hidden-project');
                }

                // Apply hidden visibility
                if (isHidden && !showHidden) {
                    folder.style.display = 'none';
                    if (hasFolderContent) folderContent.style.display = 'none';
                    return;
                }

                // Apply text filter
                if (filter) {
                    const folderMatches = projectName.includes(filter);
                    let hasMatchingFiles = false;

                    if (hasFolderContent) {
                        const fileItems = folderContent.querySelectorAll('.directory-item.file');
                        fileItems.forEach(fileItem => {
                            const fileName = (fileItem.querySelector('.filename')?.textContent || '').toLowerCase();
                            const fileMatches = folderMatches || fileName.includes(filter);
                            fileItem.style.display = fileMatches ? '' : 'none';
                            if (fileName.includes(filter)) hasMatchingFiles = true;
                        });
                    }

                    const show = folderMatches || hasMatchingFiles;
                    folder.style.display = show ? '' : 'none';
                    if (hasFolderContent) {
                        folderContent.style.display = show ? 'block' : 'none';
                    }
                } else {
                    folder.style.display = '';
                    if (hasFolderContent) {
                        const fileItems = folderContent.querySelectorAll('.directory-item.file');
                        fileItems.forEach(fi => fi.style.display = '');
                    }
                }
            });

            // Filter root-level files (not inside any folder)
            const treeChildren = document.querySelectorAll('#directoryTree > .directory-item.file');
            treeChildren.forEach(fileItem => {
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
                const fileCheckboxes = folderContent.querySelectorAll('.file-checkbox');
                fileCheckboxes.forEach(cb => {
                    cb.checked = folderCheckbox.checked;
                });
                updateBulkActionButtons();
            }
        }

        // Bulk Actions
        function updateBulkActionButtons() {
            const checkboxes = document.querySelectorAll('.file-checkbox');
            const checkedBoxes = document.querySelectorAll('.file-checkbox:checked');
            const downloadBtn = document.getElementById('bulkDownloadBtn');
            const deleteBtn = document.getElementById('bulkDeleteBtn');

            if (checkedBoxes.length > 0) {
                downloadBtn.disabled = false;
                deleteBtn.disabled = false;
            } else {
                downloadBtn.disabled = true;
                deleteBtn.disabled = true;
            }
        }

        function toggleSelectAll() {
            const checkboxes = document.querySelectorAll('.file-checkbox');
            const allChecked = Array.from(checkboxes).every(cb => cb.checked);

            checkboxes.forEach(cb => {
                cb.checked = !allChecked;
            });

            updateBulkActionButtons();
        }

        function bulkDownload() {
            const checkedBoxes = document.querySelectorAll('.file-checkbox:checked');
            if (checkedBoxes.length === 0) return;

            // Create a hidden iframe to trigger downloads
            checkedBoxes.forEach((checkbox, index) => {
                setTimeout(() => {
                    const iframe = document.createElement('iframe');
                    iframe.style.display = 'none';
                    iframe.src = 'download.php?file=' + encodeURIComponent(checkbox.dataset.file);
                    document.body.appendChild(iframe);

                    // Remove iframe after download starts
                    setTimeout(() => {
                        document.body.removeChild(iframe);
                    }, 1000);
                }, index * 500); // Stagger downloads by 500ms
            });

            showToast('success', `Downloading ${checkedBoxes.length} file(s)...`);
        }

        function bulkDelete() {
            const checkedBoxes = document.querySelectorAll('.file-checkbox:checked');
            if (checkedBoxes.length === 0) return;

            if (!confirm(`Are you sure you want to delete ${checkedBoxes.length} file(s)?`)) {
                return;
            }

            const files = Array.from(checkedBoxes).map(cb => cb.dataset.file);

            // Send delete request
            fetch('bulk_actions.php', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({
                    action: 'delete',
                    files: files
                })
            })
            .then(response => response.json())
            .then(data => {
                if (data.success) {
                    showToast('success', `Successfully deleted ${data.deleted} file(s)`);
                    setTimeout(() => location.reload(), 1500);
                } else {
                    showToast('error', data.message || 'Failed to delete files');
                }
            })
            .catch(error => {
                showToast('error', 'An error occurred while deleting files');
            });
        }

        function deleteFolder(folderPath) {
            if (!confirm(`Are you sure you want to delete the folder "${folderPath}" and all its contents?`)) {
                return;
            }

            fetch('bulk_actions.php', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({
                    action: 'delete_folder',
                    folder: folderPath
                })
            })
            .then(response => response.json())
            .then(data => {
                if (data.success) {
                    showToast('success', `Successfully deleted folder and ${data.deleted} file(s)`);
                    setTimeout(() => location.reload(), 1500);
                } else {
                    showToast('error', data.message || 'Failed to delete folder');
                }
            })
            .catch(error => {
                showToast('error', 'An error occurred while deleting folder');
            });
        }

        // Add event listeners to checkboxes and apply initial filter state
        document.addEventListener('DOMContentLoaded', function() {
            const checkboxes = document.querySelectorAll('.file-checkbox');
            checkboxes.forEach(cb => {
                cb.addEventListener('change', updateBulkActionButtons);
            });

            // Apply initial state (hide hidden projects, update button labels)
            applyFilter();
        });

        // Toast Notification System
        function showToast(type, message, options = {}) {
            const container = document.getElementById('toastContainer');
            const toast = document.createElement('div');
            toast.className = `toast toast-${type}`;

            // Determine icon based on type
            const icons = {
                success: '✓',
                error: '✕',
                processing: 'ℹ'
            };

            const titles = {
                success: 'Success',
                error: 'Error',
                processing: 'Processing'
            };

            let toastHTML = `
                <span class="toast-icon">${icons[type] || 'ℹ'}</span>
                <div class="toast-content">
                    <div class="toast-title">${titles[type] || 'Notification'}</div>
                    <div class="toast-message">${message}</div>
            `;

            // Add progress bar if provided
            if (options.processed !== undefined && options.total !== undefined) {
                const percentage = (options.processed / options.total) * 100;
                toastHTML += `
                    <div class="toast-progress">
                        <div class="toast-progress-bar" style="width: ${percentage}%"></div>
                    </div>
                `;
            }

            // Add log file link if provided
            if (options.logFile) {
                toastHTML += `
                    <div style="margin-top: 8px;">
                        <a href="${options.logFile}" target="_blank" style="color: #667eea; text-decoration: underline; font-weight: 500;">View Error Log</a>
                    </div>
                `;
            }

            toastHTML += `
                </div>
                <button class="toast-close" onclick="closeToast(this)">×</button>
            `;

            toast.innerHTML = toastHTML;
            container.appendChild(toast);

            // Auto-remove after duration
            const duration = options.duration || (type === 'error' ? 8000 : 5000);
            setTimeout(() => {
                closeToast(toast.querySelector('.toast-close'));
            }, duration);
        }

        function closeToast(button) {
            const toast = button.closest('.toast');
            toast.classList.add('hiding');
            setTimeout(() => {
                toast.remove();
            }, 300);
        }

        // Check for status on page load
        document.addEventListener('DOMContentLoaded', function() {
            const statusData = document.getElementById('statusData');
            if (statusData) {
                const type = statusData.dataset.type;
                const message = statusData.dataset.message;
                const options = {};

                if (statusData.dataset.processed && statusData.dataset.total) {
                    options.processed = parseInt(statusData.dataset.processed);
                    options.total = parseInt(statusData.dataset.total);
                }

                if (statusData.dataset.logFile) {
                    options.logFile = statusData.dataset.logFile;
                }

                showToast(type, message, options);
                statusData.remove();
            }
        });

        // Auto-refresh directory tree every 5 seconds during processing
        document.getElementById('generatorForm').addEventListener('submit', function() {
            document.getElementById('submitBtn').disabled = true;
            document.getElementById('submitBtn').textContent = 'Processing...';
            showToast('processing', 'Starting DOCX generation process...', { duration: 3000 });
        });
    </script>
</body>
</html>
