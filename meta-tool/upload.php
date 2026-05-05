<?php

declare(strict_types=1);

session_start();

require_once __DIR__ . '/lib/helper.php';

load_env(__DIR__ . '/.env');

if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
    header('Location: index.php');
    exit;
}

$target_lang = strtoupper(trim($_POST['target_lang'] ?? ''));
if (!$target_lang) {
    $_SESSION['error'] = 'Please select a target language.';
    header('Location: index.php');
    exit;
}

$upload_error = $_FILES['csv_file']['error'] ?? UPLOAD_ERR_NO_FILE;
if ($upload_error !== UPLOAD_ERR_OK) {
    $messages = [
        UPLOAD_ERR_INI_SIZE   => 'File exceeds server upload limit.',
        UPLOAD_ERR_FORM_SIZE  => 'File exceeds form size limit.',
        UPLOAD_ERR_NO_TMP_DIR => 'Server temporary directory missing.',
        UPLOAD_ERR_CANT_WRITE => 'Failed to write file to disk.',
        UPLOAD_ERR_NO_FILE    => 'No file was uploaded.',
    ];
    $_SESSION['error'] = $messages[$upload_error] ?? 'File upload failed.';
    header('Location: index.php');
    exit;
}

$file = $_FILES['csv_file'];

// Validate extension
$ext = strtolower(pathinfo($file['name'], PATHINFO_EXTENSION));
if ($ext !== 'csv') {
    $_SESSION['error'] = 'Only .csv files are accepted.';
    header('Location: index.php');
    exit;
}

// 10 MB limit
if ($file['size'] > 10 * 1024 * 1024) {
    $_SESSION['error'] = 'File is too large. Maximum allowed size is 10 MB.';
    header('Location: index.php');
    exit;
}

// Quick MIME sanity check
$finfo    = new finfo(FILEINFO_MIME_TYPE);
$mime     = $finfo->file($file['tmp_name']);
$allowed  = ['text/plain', 'text/csv', 'application/csv', 'application/octet-stream'];
if (!in_array($mime, $allowed, true)) {
    $_SESSION['error'] = 'Invalid file type. Please upload a plain CSV file.';
    header('Location: index.php');
    exit;
}

// Persist uploaded file under a random token
$tmp_dir = __DIR__ . '/tmp/';
if (!is_dir($tmp_dir)) mkdir($tmp_dir, 0755, true);

$token      = bin2hex(random_bytes(16));
$input_path = $tmp_dir . $token . '.csv';

if (!move_uploaded_file($file['tmp_name'], $input_path)) {
    $_SESSION['error'] = 'Could not save the uploaded file. Please try again.';
    header('Location: index.php');
    exit;
}

// Normalise encoding to UTF-8 (handles UTF-16 LE/BE exports from Excel)
normalize_csv_encoding($input_path);

$_SESSION['input_file']  = $input_path;
$_SESSION['target_lang'] = $target_lang;

header('Location: process.php');
exit;
