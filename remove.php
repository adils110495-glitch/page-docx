<?php
session_start();

/**
 * File removal script
 * Safely removes files from the output directory
 */

// Determine where to redirect after removal (whitelist only)
$allowed  = ['index.php', 'meta-extractor.php', 'lang-generator.php'];
$redirect = isset($_GET['redirect']) && in_array($_GET['redirect'], $allowed, true)
    ? $_GET['redirect']
    : 'index.php';

$file = isset($_GET['file']) ? $_GET['file'] : '';

if (empty($file)) {
    $_SESSION['status'] = ['type' => 'error', 'message' => 'No file specified for removal'];
    header('Location: ' . $redirect);
    exit;
}

$baseDir  = __DIR__ . '/';
$fullPath = $baseDir . $file;

// Security: must be inside output/
$realBase = realpath(__DIR__ . '/output');
$realPath = realpath($fullPath);

if ($realPath === false || strpos($realPath, $realBase) !== 0) {
    $_SESSION['status'] = ['type' => 'error', 'message' => 'Invalid file path'];
    header('Location: ' . $redirect);
    exit;
}

if (!file_exists($fullPath)) {
    $_SESSION['status'] = ['type' => 'error', 'message' => 'File not found'];
    header('Location: ' . $redirect);
    exit;
}

if (!is_file($fullPath)) {
    $_SESSION['status'] = ['type' => 'error', 'message' => 'Cannot remove directories'];
    header('Location: ' . $redirect);
    exit;
}

$sessionKey = ($redirect === 'lang-generator.php') ? 'lang_status' : 'status';

if (unlink($fullPath)) {
    $_SESSION[$sessionKey] = ['type' => 'success', 'message' => 'File deleted successfully: ' . basename($fullPath)];
} else {
    $_SESSION[$sessionKey] = ['type' => 'error', 'message' => 'Failed to delete file. Check permissions.'];
}

header('Location: ' . $redirect);
exit;
