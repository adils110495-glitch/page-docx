<?php
session_start();

/**
 * Bulk Actions Handler
 * Handles bulk download and delete operations
 */

// Get JSON input
$input = file_get_contents('php://input');
$data = json_decode($input, true);

if (!$data || !isset($data['action']) || !isset($data['files'])) {
    http_response_code(400);
    echo json_encode(['success' => false, 'message' => 'Invalid request']);
    exit;
}

$action = $data['action'];
$files = $data['files'];

// Security: Validate all files are in output directory
$baseDir = __DIR__ . '/output';
$validFiles = [];

foreach ($files as $file) {
    // Remove 'output/' prefix if present
    $cleanFile = str_replace('output/', '', $file);
    $fullPath = $baseDir . '/' . $cleanFile;

    // Security checks
    if (strpos($cleanFile, '..') !== false) {
        continue; // Skip directory traversal attempts
    }

    if (!file_exists($fullPath) || !is_file($fullPath)) {
        continue; // Skip non-existent or non-file paths
    }

    // Verify real path is within output directory
    $realBase = realpath($baseDir);
    $realPath = realpath($fullPath);

    if ($realPath === false || strpos($realPath, $realBase) !== 0) {
        continue; // Skip files outside output directory
    }

    $validFiles[] = $realPath;
}

if (empty($validFiles)) {
    echo json_encode(['success' => false, 'message' => 'No valid files to process']);
    exit;
}

// Handle action
if ($action === 'delete') {
    $deletedCount = 0;
    $errors = [];

    foreach ($validFiles as $filePath) {
        if (unlink($filePath)) {
            $deletedCount++;
        } else {
            $errors[] = basename($filePath);
        }
    }

    if ($deletedCount > 0) {
        echo json_encode([
            'success' => true,
            'deleted' => $deletedCount,
            'errors' => $errors
        ]);
    } else {
        echo json_encode([
            'success' => false,
            'message' => 'Failed to delete any files'
        ]);
    }
} else {
    http_response_code(400);
    echo json_encode(['success' => false, 'message' => 'Unknown action']);
}
