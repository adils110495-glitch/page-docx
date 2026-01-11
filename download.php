<?php
session_start();

/**
 * File download script
 * Safely serves DOCX files from the output directory with proper headers
 */

// Get the file path from query parameter
$file = isset($_GET['file']) ? $_GET['file'] : '';

if (empty($file)) {
    header('HTTP/1.0 404 Not Found');
    echo 'File not specified';
    exit;
}

// Build the full path
$baseDir = __DIR__ . '/';
$fullPath = $baseDir . $file;

// Security check: ensure the file is within the output directory
// Normalize the path and check it starts with 'output/'
$normalizedFile = str_replace('\\', '/', $file);
if (strpos($normalizedFile, 'output/') !== 0) {
    header('HTTP/1.0 403 Forbidden');
    echo 'Access denied - Invalid path';
    exit;
}

// Additional security: prevent directory traversal
if (strpos($normalizedFile, '..') !== false) {
    header('HTTP/1.0 403 Forbidden');
    echo 'Access denied - Directory traversal detected';
    exit;
}

// Verify the real path is within output directory after file exists
if (file_exists($fullPath)) {
    $realBase = realpath(__DIR__ . '/output');
    $realPath = realpath($fullPath);

    if ($realPath === false || strpos($realPath, $realBase) !== 0) {
        header('HTTP/1.0 403 Forbidden');
        echo 'Access denied - Path validation failed';
        exit;
    }
}

// Check if file exists
if (!file_exists($fullPath)) {
    header('HTTP/1.0 404 Not Found');
    echo 'File not found';
    exit;
}

// Check if it's a file (not a directory)
if (!is_file($fullPath)) {
    header('HTTP/1.0 403 Forbidden');
    echo 'Invalid file';
    exit;
}

// Only allow DOCX files
$extension = strtolower(pathinfo($fullPath, PATHINFO_EXTENSION));
if ($extension !== 'docx') {
    header('HTTP/1.0 403 Forbidden');
    echo 'Only DOCX files can be downloaded';
    exit;
}

// Get file info
$filename = basename($fullPath);
$filesize = filesize($fullPath);

// Set headers for download
header('Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document');
header('Content-Disposition: attachment; filename="' . $filename . '"');
header('Content-Length: ' . $filesize);
header('Cache-Control: no-cache, must-revalidate');
header('Expires: 0');
header('Pragma: public');

// Clear output buffer
if (ob_get_level()) {
    ob_clean();
}

// Read and output file
readfile($fullPath);
exit;
