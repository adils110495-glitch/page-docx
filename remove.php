<?php
session_start();

/**
 * File removal script
 * Safely removes files from the output directory
 */

// Get the file path from query parameter
$file = isset($_GET['file']) ? $_GET['file'] : '';

if (empty($file)) {
    $_SESSION['status'] = [
        'type' => 'error',
        'message' => 'No file specified for removal'
    ];
    header('Location: index.php');
    exit;
}

// Build the full path
$baseDir = __DIR__ . '/';
$fullPath = $baseDir . $file;

// Security check: ensure the file is within the output directory
$realBase = realpath(__DIR__ . '/output');
$realPath = realpath($fullPath);

if ($realPath === false || strpos($realPath, $realBase) !== 0) {
    $_SESSION['status'] = [
        'type' => 'error',
        'message' => 'Invalid file path'
    ];
    header('Location: index.php');
    exit;
}

// Check if file exists
if (!file_exists($fullPath)) {
    $_SESSION['status'] = [
        'type' => 'error',
        'message' => 'File not found'
    ];
    header('Location: index.php');
    exit;
}

// Check if it's a file (not a directory)
if (!is_file($fullPath)) {
    $_SESSION['status'] = [
        'type' => 'error',
        'message' => 'Cannot remove directories'
    ];
    header('Location: index.php');
    exit;
}

// Attempt to delete the file
if (unlink($fullPath)) {
    $filename = basename($fullPath);
    $_SESSION['status'] = [
        'type' => 'success',
        'message' => 'File deleted successfully: ' . $filename
    ];
} else {
    $_SESSION['status'] = [
        'type' => 'error',
        'message' => 'Failed to delete file. Check permissions.'
    ];
}

// Redirect back to index
header('Location: index.php');
exit;
