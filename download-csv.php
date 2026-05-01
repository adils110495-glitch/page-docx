<?php
/**
 * CSV file download script
 * Safely serves CSV files from the output directory
 */

$file = isset($_GET['file']) ? $_GET['file'] : '';

if (empty($file)) {
    header('HTTP/1.0 404 Not Found');
    echo 'File not specified';
    exit;
}

$baseDir  = __DIR__ . '/';
$fullPath = $baseDir . $file;

// Must be inside output/
$normalizedFile = str_replace('\\', '/', $file);
if (strpos($normalizedFile, 'output/') !== 0) {
    header('HTTP/1.0 403 Forbidden');
    echo 'Access denied - Invalid path';
    exit;
}

// No directory traversal
if (strpos($normalizedFile, '..') !== false) {
    header('HTTP/1.0 403 Forbidden');
    echo 'Access denied - Directory traversal detected';
    exit;
}

// Verify real path stays inside output/
if (file_exists($fullPath)) {
    $realBase = realpath(__DIR__ . '/output');
    $realPath = realpath($fullPath);

    if ($realPath === false || strpos($realPath, $realBase) !== 0) {
        header('HTTP/1.0 403 Forbidden');
        echo 'Access denied - Path validation failed';
        exit;
    }
}

if (!file_exists($fullPath)) {
    header('HTTP/1.0 404 Not Found');
    echo 'File not found';
    exit;
}

if (!is_file($fullPath)) {
    header('HTTP/1.0 403 Forbidden');
    echo 'Invalid file';
    exit;
}

// Only allow CSV files
$extension = strtolower(pathinfo($fullPath, PATHINFO_EXTENSION));
if ($extension !== 'csv') {
    header('HTTP/1.0 403 Forbidden');
    echo 'Only CSV files can be downloaded here';
    exit;
}

$filename = basename($fullPath);
$filesize = filesize($fullPath);

header('Content-Type: text/csv; charset=UTF-8');
header('Content-Disposition: attachment; filename="' . $filename . '"');
header('Content-Length: ' . $filesize);
header('Cache-Control: no-cache, must-revalidate');
header('Expires: 0');
header('Pragma: public');

if (ob_get_level()) {
    ob_clean();
}

readfile($fullPath);
exit;
